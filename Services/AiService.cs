using System.Diagnostics;
using System.IO.Compression;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;
using TaskTool.Models;

namespace TaskTool.Services;

public interface IAiService { Task<string> SendAsync(string systemPrompt, string userPrompt, CancellationToken cancellationToken = default); }

public sealed class OpenAiCompatibleAiProvider : IAiService
{
    private readonly HttpClient _httpClient; private readonly Uri _endpoint; private readonly string _model; private readonly string _apiKey;
    public OpenAiCompatibleAiProvider(HttpClient httpClient, string baseUrl, string model, string apiKey = "")
    { _httpClient = httpClient; _endpoint = BuildEndpoint(baseUrl); _model = string.IsNullOrWhiteSpace(model) ? throw new ArgumentException("Ein Modellname ist erforderlich.", nameof(model)) : model.Trim(); _apiKey = apiKey; }
    public async Task<string> SendAsync(string systemPrompt, string userPrompt, CancellationToken cancellationToken = default)
    {
        using var request = new HttpRequestMessage(HttpMethod.Post, _endpoint) { Content = JsonContent.Create(new { model = _model, messages = new[] { new { role = "system", content = systemPrompt }, new { role = "user", content = userPrompt } }, temperature = 0, max_tokens = 32 }) };
        if (!string.IsNullOrWhiteSpace(_apiKey)) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
        using var response = await _httpClient.SendAsync(request, cancellationToken); response.EnsureSuccessStatusCode();
        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken); using var document = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
        if (!document.RootElement.TryGetProperty("choices", out var choices) || choices.GetArrayLength() == 0 || !choices[0].TryGetProperty("message", out var message) || !message.TryGetProperty("content", out var content)) throw new InvalidDataException("Die KI-Antwort enthält keinen Text im erwarteten Chat-Completions-Format.");
        return content.GetString()?.Trim() ?? string.Empty;
    }
    internal static Uri BuildEndpoint(string baseUrl)
    {
        if (!Uri.TryCreate(baseUrl?.Trim(), UriKind.Absolute, out var uri) || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps)) throw new UriFormatException("Die API Basis-URL muss eine absolute HTTP- oder HTTPS-URL sein.");
        var root = uri.ToString().TrimEnd('/'); return new Uri(root.EndsWith("/v1", StringComparison.OrdinalIgnoreCase) ? root + "/chat/completions" : root + "/v1/chat/completions");
    }
}

public enum LocalAiStatus { NotInstalled, DownloadingRuntime, DownloadingModel, VerifyingSha256, Installed, LoadingModel, Ready, Error }

public sealed class AiService : IDisposable
{
    public const string TestSystemPrompt = "Folge der Benutzeranweisung exakt. Gib keine zusätzlichen Erklärungen aus.";
    public const string TestUserPrompt = "Antworte ausschließlich mit exakt: Test erfolgreich";
    private readonly SettingsService _settings; private readonly HttpClient _httpClient;
    public AiService(SettingsService settings, LoggerService logger, HttpClient? httpClient = null) { _settings = settings; _httpClient = httpClient ?? new HttpClient { Timeout = TimeSpan.FromMinutes(10) }; LocalServer = new LocalLlamaServerManager(settings, logger, _httpClient); }
    public LocalLlamaServerManager LocalServer { get; }
    public IAiService CreateProvider() => _settings.Current.AiProvider switch
    {
        AiProviderType.OpenAiCompatible => new OpenAiCompatibleAiProvider(_httpClient, _settings.Current.AiApiBaseUrl, _settings.Current.AiModel, _settings.GetAiApiKey()),
        AiProviderType.LocalLlama when LocalServer.IsReady => new OpenAiCompatibleAiProvider(_httpClient, LocalServer.ApiBaseUrl!, LocalAiModelCatalog.Get(_settings.Current.AiLocalPreset).LlamaAlias),
        AiProviderType.LocalLlama => throw new InvalidOperationException("Die lokale KI ist noch nicht bereit."),
        _ => throw new InvalidOperationException("Der gewählte KI-Provider wird nicht unterstützt.")
    };
    public async Task<string> TestAsync(CancellationToken cancellationToken = default)
    {
        if (!_settings.Current.AiEnabled) throw new InvalidOperationException("KI ist deaktiviert.");
        var answer = await CreateProvider().SendAsync(TestSystemPrompt, TestUserPrompt, cancellationToken);
        return Regex.Replace(answer, @"<think>[\s\S]*?</think>", "", RegexOptions.IgnoreCase).Trim();
    }
    public async Task InitializeLocalInBackgroundAsync(CancellationToken token = default)
    { if (_settings.Current.AiEnabled && _settings.Current.AiProvider == AiProviderType.LocalLlama) await LocalServer.InstallAndStartAsync(null, token); }
    public void Dispose() { LocalServer.Dispose(); _httpClient.Dispose(); }
}

public sealed class LocalLlamaServerManager : IDisposable
{
    public const string Host = "127.0.0.1";
    private const string ReleasesApi = "https://api.github.com/repos/ggml-org/llama.cpp/releases/latest";
    private readonly SettingsService _settings; private readonly LoggerService _logger; private readonly HttpClient _httpClient; private readonly SemaphoreSlim _gate = new(1, 1);
    private readonly Queue<string> _diagnostics = new(); private Process? _process; private int _port;
    public LocalLlamaServerManager(SettingsService settings, LoggerService logger, HttpClient httpClient) { _settings = settings; _logger = logger; _httpClient = httpClient; }
    public string RootDirectory { get; } = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Plenaro", "AI");
    public string RuntimeDirectory => Path.Combine(RootDirectory, "runtime", "llama.cpp");
    public string RuntimeExecutable => Path.Combine(RuntimeDirectory, "llama-server.exe");
    public string ModelPath(LocalAiPreset preset) => Path.Combine(RootDirectory, "models", preset.ToString().ToLowerInvariant(), LocalAiModelCatalog.Get(preset).FileName);
    public bool IsRunning => _process is { HasExited: false }; public bool IsReady { get; private set; }
    public string? ApiBaseUrl => _port == 0 ? null : $"http://{Host}:{_port}/v1";
    public LocalAiStatus Status { get; private set; } = LocalAiStatus.NotInstalled; public int Progress { get; private set; }
    public event EventHandler? StateChanged;
    private void SetStatus(LocalAiStatus status, int progress = 0) { Status = status; Progress = progress; StateChanged?.Invoke(this, EventArgs.Empty); }

    public async Task InstallAndStartAsync(IProgress<int>? progress, CancellationToken token = default)
    {
        await _gate.WaitAsync(token);
        try
        {
            if (!_settings.Current.AiEnabled || _settings.Current.AiProvider != AiProviderType.LocalLlama) return;
            var preset = _settings.Current.AiLocalPreset; Stop();
            if (!File.Exists(RuntimeExecutable)) await DownloadRuntimeAsync(progress, token);
            if (!await IsModelValidAsync(preset, token))
            {
                var invalidModel = ModelPath(preset);
                if (File.Exists(invalidModel)) File.Delete(invalidModel);
                await DownloadModelAsync(preset, progress, token);
            }
            SetStatus(LocalAiStatus.Installed); await StartAndWaitForReadyAsync(preset, token);
        }
        catch { SetStatus(LocalAiStatus.Error); throw; }
        finally { _gate.Release(); }
    }

    public async Task DownloadModelAsync(LocalAiPreset preset, IProgress<int>? progress = null, CancellationToken token = default)
    {
        var model = LocalAiModelCatalog.Get(preset); var uri = new Uri(model.DownloadUrl);
        if (uri.Scheme != Uri.UriSchemeHttps) throw new InvalidOperationException("Modelle dürfen nur über HTTPS geladen werden.");
        var destination = ModelPath(preset); var part = destination + ".part"; Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
        try { SetStatus(LocalAiStatus.DownloadingModel); await DownloadAsync(uri, part, progress, token); SetStatus(LocalAiStatus.VerifyingSha256); if (!await HasSha256Async(part, model.Sha256, token)) throw new InvalidDataException("Die SHA256-Prüfsumme des Modells stimmt nicht überein."); File.Move(part, destination, true); }
        finally { if (File.Exists(part)) File.Delete(part); }
    }
    public async Task<bool> IsModelValidAsync(LocalAiPreset preset, CancellationToken token = default) { var path = ModelPath(preset); return File.Exists(path) && !File.Exists(path + ".part") && await HasSha256Async(path, LocalAiModelCatalog.Get(preset).Sha256, token); }

    private async Task DownloadRuntimeAsync(IProgress<int>? progress, CancellationToken token)
    {
        SetStatus(LocalAiStatus.DownloadingRuntime); using var request = new HttpRequestMessage(HttpMethod.Get, ReleasesApi); request.Headers.UserAgent.ParseAdd("Plenaro/1.0");
        using var response = await _httpClient.SendAsync(request, token); response.EnsureSuccessStatusCode(); using var json = JsonDocument.Parse(await response.Content.ReadAsStringAsync(token));
        var asset = json.RootElement.GetProperty("assets").EnumerateArray().FirstOrDefault(a => Regex.IsMatch(a.GetProperty("name").GetString() ?? "", @"^llama-.*-bin-win-cpu-x64\.zip$", RegexOptions.IgnoreCase));
        if (asset.ValueKind == JsonValueKind.Undefined) throw new InvalidDataException("Kein offizielles Windows-x64-CPU-Asset gefunden.");
        var url = new Uri(asset.GetProperty("browser_download_url").GetString()!); if (url.Scheme != Uri.UriSchemeHttps || url.Host != "github.com") throw new InvalidDataException("Ungültige Runtime-Quelle.");
        Directory.CreateDirectory(RootDirectory); var part = Path.Combine(RootDirectory, "llama-runtime.zip.part");
        try { await DownloadAsync(url, part, progress, token); if (asset.TryGetProperty("digest", out var digest) && digest.GetString() is { } value && value.StartsWith("sha256:", StringComparison.OrdinalIgnoreCase) && !await HasSha256Async(part, value[7..], token)) throw new InvalidDataException("Die SHA256-Prüfsumme der Runtime stimmt nicht überein."); ExtractZipSafely(part, RuntimeDirectory); NormalizeRuntimeLayout(); if (!File.Exists(RuntimeExecutable)) throw new InvalidDataException("llama-server.exe fehlt im Runtime-Archiv."); }
        finally { if (File.Exists(part)) File.Delete(part); }
    }
    private async Task DownloadAsync(Uri uri, string target, IProgress<int>? progress, CancellationToken token)
    {
        using var response = await _httpClient.GetAsync(uri, HttpCompletionOption.ResponseHeadersRead, token); response.EnsureSuccessStatusCode(); var length = response.Content.Headers.ContentLength;
        await using var input = await response.Content.ReadAsStreamAsync(token); await using var output = new FileStream(target, FileMode.Create, FileAccess.Write, FileShare.None, 131072, true); var buffer = new byte[131072]; long total = 0;
        for (var read = await input.ReadAsync(buffer, token); read > 0; read = await input.ReadAsync(buffer, token)) { await output.WriteAsync(buffer.AsMemory(0, read), token); total += read; var percent = length > 0 ? (int)(total * 100 / length) : 0; Progress = percent; progress?.Report(percent); StateChanged?.Invoke(this, EventArgs.Empty); } await output.FlushAsync(token);
    }
    internal static void ExtractZipSafely(string archive, string destination)
    {
        Directory.CreateDirectory(destination); var root = Path.GetFullPath(destination) + Path.DirectorySeparatorChar; using var zip = ZipFile.OpenRead(archive);
        foreach (var entry in zip.Entries) { var path = Path.GetFullPath(Path.Combine(destination, entry.FullName)); if (!path.StartsWith(root, StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Unsicherer Pfad im Runtime-Archiv."); if (string.IsNullOrEmpty(entry.Name)) Directory.CreateDirectory(path); else { Directory.CreateDirectory(Path.GetDirectoryName(path)!); entry.ExtractToFile(path, true); } }
    }
    private void NormalizeRuntimeLayout()
    {
        if (File.Exists(RuntimeExecutable)) return;
        var executable = Directory.EnumerateFiles(RuntimeDirectory, "llama-server.exe", SearchOption.AllDirectories).FirstOrDefault();
        if (executable == null) return;
        var source = Path.GetDirectoryName(executable)!;
        foreach (var file in Directory.EnumerateFiles(source)) File.Move(file, Path.Combine(RuntimeDirectory, Path.GetFileName(file)), true);
    }
    internal static async Task<bool> HasSha256Async(string path, string expected, CancellationToken token = default) { await using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 131072, true); var hash = await SHA256.HashDataAsync(stream, token); return Convert.ToHexString(hash).Equals(expected, StringComparison.OrdinalIgnoreCase); }

    private async Task StartAndWaitForReadyAsync(LocalAiPreset preset, CancellationToken token)
    {
        _port = GetFreePort(); var info = new ProcessStartInfo(RuntimeExecutable) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
        foreach (var arg in new[] { "-m", ModelPath(preset), "--host", Host, "--port", _port.ToString(), "--ctx-size", "4096", "--parallel", "1", "--alias", LocalAiModelCatalog.Get(preset).LlamaAlias }) info.ArgumentList.Add(arg);
        _process = new Process { StartInfo = info, EnableRaisingEvents = true }; _process.OutputDataReceived += CaptureDiagnostic; _process.ErrorDataReceived += CaptureDiagnostic;
        if (!_process.Start()) throw new InvalidOperationException("llama.cpp konnte nicht gestartet werden."); _process.BeginOutputReadLine(); _process.BeginErrorReadLine(); SetStatus(LocalAiStatus.LoadingModel); _logger.Info($"[AI] Local llama.cpp server started host={Host} port={_port}.");
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(token); timeout.CancelAfter(TimeSpan.FromMinutes(3));
        while (!timeout.IsCancellationRequested) { if (_process.HasExited) throw new InvalidOperationException("llama.cpp wurde während des Ladens beendet. " + string.Join(" | ", _diagnostics)); try { using var response = await _httpClient.GetAsync($"http://{Host}:{_port}/health", timeout.Token); if (response.IsSuccessStatusCode) { IsReady = true; SetStatus(LocalAiStatus.Ready, 100); return; } } catch (HttpRequestException) { } await Task.Delay(500, timeout.Token); }
        throw new TimeoutException("llama.cpp wurde nicht innerhalb von drei Minuten bereit.");
    }
    private void CaptureDiagnostic(object sender, DataReceivedEventArgs e) { if (string.IsNullOrWhiteSpace(e.Data)) return; lock (_diagnostics) { _diagnostics.Enqueue(e.Data); while (_diagnostics.Count > 30) _diagnostics.Dequeue(); } }
    internal static int GetFreePort() { var listener = new TcpListener(IPAddress.Loopback, 0); listener.Start(); var port = ((IPEndPoint)listener.LocalEndpoint).Port; listener.Stop(); return port; }
    public void Stop() { IsReady = false; if (_process is { HasExited: false }) { _process.Kill(entireProcessTree: true); } _process?.Dispose(); _process = null; _port = 0; StateChanged?.Invoke(this, EventArgs.Empty); }
    public void Dispose() { Stop(); _gate.Dispose(); }
}
