using System.Diagnostics;
using System.IO;
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
    private readonly Queue<string> _diagnostics = new(); private Process? _process; private int _port; private string _setupStage = "initialization";
    public LocalLlamaServerManager(SettingsService settings, LoggerService logger, HttpClient httpClient) { _settings = settings; _logger = logger; _httpClient = httpClient; }
    public string RootDirectory { get; } = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Plenaro", "AI");
    public string RuntimeDirectory => Path.Combine(RootDirectory, "runtime", "llama.cpp");
    public string RuntimeExecutable => Path.Combine(RuntimeDirectory, "llama-server.exe");
    public string ModelPath(LocalAiPreset preset) => Path.Combine(RootDirectory, "models", preset.ToString().ToLowerInvariant(), LocalAiModelCatalog.Get(preset).FileName);
    public bool IsRunning => _process is { HasExited: false }; public bool IsReady { get; private set; }
    public string? ApiBaseUrl => _port == 0 ? null : $"http://{Host}:{_port}/v1";
    public LocalAiStatus Status { get; private set; } = LocalAiStatus.NotInstalled; public int Progress { get; private set; }
    public string? LastError { get; private set; }
    public event EventHandler? StateChanged;
    private void SetStatus(LocalAiStatus status, int progress = 0) { Status = status; Progress = progress; StateChanged?.Invoke(this, EventArgs.Empty); }

    public async Task InstallAndStartAsync(IProgress<int>? progress, CancellationToken token = default)
    {
        await _gate.WaitAsync(token);
        try
        {
            if (!_settings.Current.AiEnabled || _settings.Current.AiProvider != AiProviderType.LocalLlama) return;
            LastError = null;
            var preset = _settings.Current.AiLocalPreset;
            Log($"Local AI setup started preset={preset}");
            Stop();
            preset = await EnsureInstalledAsync(progress, token);
            _setupStage = "server-start";
            await StartAndWaitForReadyAsync(preset, token);
            Log("Local AI setup completed");
        }
        catch (Exception exception) { HandleSetupFailure(exception); throw; }
        finally { _gate.Release(); }
    }

    public async Task InstallAsync(IProgress<int>? progress, CancellationToken token = default)
    {
        await _gate.WaitAsync(token);
        try
        {
            if (!_settings.Current.AiEnabled || _settings.Current.AiProvider != AiProviderType.LocalLlama) return;
            LastError = null;
            Log($"Local AI setup started preset={_settings.Current.AiLocalPreset}");
            await EnsureInstalledAsync(progress, token);
            Log("Local AI setup completed");
        }
        catch (Exception exception) { HandleSetupFailure(exception); throw; }
        finally { _gate.Release(); }
    }

    private async Task<LocalAiPreset> EnsureInstalledAsync(IProgress<int>? progress, CancellationToken token)
    {
        var preset = _settings.Current.AiLocalPreset;
        _setupStage = "runtime-check";
        Log("Checking llama.cpp runtime");
        if (!File.Exists(RuntimeExecutable)) await DownloadRuntimeAsync(progress, token);
        _setupStage = "model-check";
        Log($"Checking model preset={preset}");
        if (!await IsModelValidAsync(preset, token))
        {
            Log("Model download started");
            var invalidModel = ModelPath(preset);
            if (File.Exists(invalidModel)) File.Delete(invalidModel);
            await DownloadModelAsync(preset, progress, token);
        }
        else Log("Model already installed");
        SetStatus(LocalAiStatus.Installed, 100);
        return preset;
    }

    public async Task DownloadModelAsync(LocalAiPreset preset, IProgress<int>? progress = null, CancellationToken token = default)
    {
        _setupStage = "model-download";
        var model = LocalAiModelCatalog.Get(preset); var uri = new Uri(model.DownloadUrl);
        if (uri.Scheme != Uri.UriSchemeHttps) throw new InvalidOperationException("Modelle dürfen nur über HTTPS geladen werden.");
        var destination = ModelPath(preset); var part = destination + ".part"; Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
        try
        {
            SetStatus(LocalAiStatus.DownloadingModel); await DownloadAsync(uri, part, progress, token); SetStatus(LocalAiStatus.VerifyingSha256);
            await PromoteVerifiedDownloadAsync(part, destination, model.Sha256, token);
            Log("Model SHA256 verified");
        }
        finally { if (File.Exists(part)) File.Delete(part); }
    }
    public Task<bool> IsModelValidAsync(LocalAiPreset preset, CancellationToken token = default) => IsModelFileValidAsync(ModelPath(preset), LocalAiModelCatalog.Get(preset).Sha256, token);
    internal static async Task<bool> IsModelFileValidAsync(string path, string expectedSha256, CancellationToken token = default) => File.Exists(path) && !File.Exists(path + ".part") && await HasSha256Async(path, expectedSha256, token);

    internal static async Task PromoteVerifiedDownloadAsync(string partPath, string destination, string expectedSha256, CancellationToken token = default)
    {
        try
        {
            if (!await HasSha256Async(partPath, expectedSha256, token)) throw new InvalidDataException("Die SHA256-Prüfsumme des Modells stimmt nicht überein.");
            File.Move(partPath, destination, true);
        }
        finally { if (File.Exists(partPath)) File.Delete(partPath); }
    }

    private async Task DownloadRuntimeAsync(IProgress<int>? progress, CancellationToken token)
    {
        _setupStage = "runtime-metadata";
        SetStatus(LocalAiStatus.DownloadingRuntime);
        Log("Fetching llama.cpp release metadata");
        using var request = new HttpRequestMessage(HttpMethod.Get, ReleasesApi); request.Headers.UserAgent.ParseAdd("Plenaro/1.0");
        using var response = await _httpClient.SendAsync(request, token);
        if (!response.IsSuccessStatusCode) throw new HttpRequestException($"llama.cpp-Release-Metadaten konnten nicht geladen werden (HTTP {(int)response.StatusCode}).", null, response.StatusCode);
        await using var responseStream = await response.Content.ReadAsStreamAsync(token);
        using var json = await JsonDocument.ParseAsync(responseStream, cancellationToken: token);
        if (!json.RootElement.TryGetProperty("assets", out var assetsElement) || assetsElement.ValueKind != JsonValueKind.Array)
            throw new InvalidDataException("Die llama.cpp-Release-Antwort enthält keine gültige Assetliste.");
        var assets = assetsElement.EnumerateArray()
            .Where(item => item.ValueKind == JsonValueKind.Object)
            .Select(item => new { Json = item, Name = item.TryGetProperty("name", out var name) && name.ValueKind == JsonValueKind.String ? name.GetString() : null })
            .ToList();
        string selectedName;
        try { selectedName = SelectWindowsX64CpuAsset(assets.Select(item => item.Name)); }
        catch (InvalidDataException)
        {
            var available = assets.Select(item => item.Name).Where(name => !string.IsNullOrWhiteSpace(name));
            _logger.Error($"[AI] llama.cpp Windows x64 CPU runtime asset not found. Available assets: {string.Join(", ", available)}");
            throw;
        }
        var asset = assets.First(item => string.Equals(item.Name, selectedName, StringComparison.OrdinalIgnoreCase)).Json;
        Log($"Selected runtime asset={selectedName}");
        if (!asset.TryGetProperty("browser_download_url", out var urlElement) || urlElement.ValueKind != JsonValueKind.String ||
            !Uri.TryCreate(urlElement.GetString(), UriKind.Absolute, out var url) || url.Scheme != Uri.UriSchemeHttps ||
            !string.Equals(url.Host, "github.com", StringComparison.OrdinalIgnoreCase) ||
            !url.AbsolutePath.StartsWith("/ggml-org/llama.cpp/releases/download/", StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("Das Runtime-Asset verweist nicht auf die erwartete HTTPS-GitHub-Quelle.");
        _setupStage = "runtime-download";
        Directory.CreateDirectory(RootDirectory); var part = Path.Combine(RootDirectory, "llama-runtime.zip.part");
        try { Log("Downloading llama.cpp runtime"); await DownloadAsync(url, part, progress, token); Log("Runtime download completed"); if (asset.TryGetProperty("digest", out var digest) && digest.ValueKind == JsonValueKind.String && digest.GetString() is { } value && value.StartsWith("sha256:", StringComparison.OrdinalIgnoreCase)) { if (!await HasSha256Async(part, value[7..], token)) throw new InvalidDataException("Die SHA256-Prüfsumme der Runtime stimmt nicht überein."); Log("Runtime SHA256 verified"); } ExtractZipSafely(part, RuntimeDirectory); NormalizeRuntimeLayout(); if (!File.Exists(RuntimeExecutable)) throw new InvalidDataException("llama-server.exe fehlt im Runtime-Archiv."); Log("Runtime extracted"); }
        finally { if (File.Exists(part)) File.Delete(part); }
    }

    internal static string SelectWindowsX64CpuAsset(IEnumerable<string?> assetNames)
    {
        const string officialAssetName = "llama-bin-win-cpu-x64.zip";
        var match = assetNames.FirstOrDefault(name => string.Equals(name, officialAssetName, StringComparison.OrdinalIgnoreCase));
        return match ?? throw new InvalidDataException("Kein offizielles Windows-x64-CPU-Asset gefunden.");
    }

    private void Log(string message) => _logger.OperationalInfo($"[AI] {message}");
    private void HandleSetupFailure(Exception exception)
    {
        LastError = exception.Message;
        SetStatus(LocalAiStatus.Error);
        _logger.Error($"[AI] Local AI setup failed stage={_setupStage}: {exception.GetType().Name}: {exception.Message}");
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
