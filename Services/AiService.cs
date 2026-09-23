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

public interface IAiService
{
    Task<string> ChatAsync(IReadOnlyList<AiChatRequestMessage> messages, AiRequestOptions options, CancellationToken cancellationToken = default);
    Task<string> SendAsync(string systemPrompt, string userPrompt, CancellationToken cancellationToken = default);
}

public interface IAiChatService
{
    bool IsEnabled { get; }
    bool CanChat { get; }
    string ProviderDescription { get; }
    string AvailabilityMessage { get; }
    event EventHandler? StateChanged;
    Task<string> ChatAsync(IReadOnlyList<AiChatRequestMessage> messages, AiRequestOptions options, CancellationToken cancellationToken = default);
}

public sealed class OpenAiCompatibleAiProvider : IAiService
{
    private readonly HttpClient _httpClient; private readonly Uri _endpoint; private readonly string _model; private readonly string _apiKey;
    public OpenAiCompatibleAiProvider(HttpClient httpClient, string baseUrl, string model, string apiKey = "")
    { _httpClient = httpClient; _endpoint = BuildEndpoint(baseUrl); _model = string.IsNullOrWhiteSpace(model) ? throw new ArgumentException("Ein Modellname ist erforderlich.", nameof(model)) : model.Trim(); _apiKey = apiKey; }
    public async Task<string> SendAsync(string systemPrompt, string userPrompt, CancellationToken cancellationToken = default)
        => await ChatAsync(
            [new(AiChatRole.System, systemPrompt), new(AiChatRole.User, userPrompt)],
            new AiRequestOptions(0, 32), cancellationToken);

    public async Task<string> ChatAsync(IReadOnlyList<AiChatRequestMessage> messages, AiRequestOptions options, CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(messages);
        if (messages.Count == 0) throw new ArgumentException("Mindestens eine Chat-Nachricht ist erforderlich.", nameof(messages));
        if (options.MaxTokens <= 0) throw new ArgumentOutOfRangeException(nameof(options));
        var requestMessages = messages.Select(message => new { role = message.Role.ToString().ToLowerInvariant(), content = message.Content }).ToArray();
        using var request = new HttpRequestMessage(HttpMethod.Post, _endpoint) { Content = JsonContent.Create(new { model = _model, messages = requestMessages, temperature = options.Temperature, max_tokens = options.MaxTokens }) };
        if (!string.IsNullOrWhiteSpace(_apiKey)) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
        using var response = await _httpClient.SendAsync(request, cancellationToken); response.EnsureSuccessStatusCode();
        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken); using var document = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
        if (!document.RootElement.TryGetProperty("choices", out var choices) || choices.GetArrayLength() == 0 || !choices[0].TryGetProperty("message", out var message) || !message.TryGetProperty("content", out var content)) throw new InvalidDataException("Die KI-Antwort enthält keinen Text im erwarteten Chat-Completions-Format.");
        var answer = content.GetString()?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(answer)) throw new InvalidDataException("Die KI-Antwort ist leer.");
        return answer;
    }
    internal static Uri BuildEndpoint(string baseUrl)
    {
        if (!Uri.TryCreate(baseUrl?.Trim(), UriKind.Absolute, out var uri) || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps)) throw new UriFormatException("Die API Basis-URL muss eine absolute HTTP- oder HTTPS-URL sein.");
        var root = uri.ToString().TrimEnd('/'); return new Uri(root.EndsWith("/v1", StringComparison.OrdinalIgnoreCase) ? root + "/chat/completions" : root + "/v1/chat/completions");
    }
}

public enum LocalAiStatus { NotInstalled, DownloadingRuntime, DownloadingModel, VerifyingSha256, Installed, LoadingModel, Ready, Error }

public sealed class AiService : IAiChatService, IDisposable
{
    public const string TestSystemPrompt = "Folge der Benutzeranweisung exakt. Gib keine zusätzlichen Erklärungen aus.";
    public const string TestUserPrompt = "Antworte ausschließlich mit exakt: Test erfolgreich";
    private readonly SettingsService _settings; private readonly HttpClient _httpClient; private readonly LoggerService _logger;
    public AiService(SettingsService settings, LoggerService logger, HttpClient? httpClient = null) { _settings = settings; _logger = logger; _httpClient = httpClient ?? new HttpClient { Timeout = TimeSpan.FromMinutes(10) }; LocalServer = new LocalLlamaServerManager(settings, logger, _httpClient); LocalServer.StateChanged += OnStateChanged; _settings.SettingsChanged += OnSettingsChanged; }
    public LocalLlamaServerManager LocalServer { get; }
    public bool IsEnabled => _settings.Current.AiEnabled;
    public bool CanChat => IsEnabled && (_settings.Current.AiProvider != AiProviderType.LocalLlama || LocalServer.IsReady);
    public string ProviderDescription => _settings.Current.AiProvider == AiProviderType.LocalLlama
        ? $"Lokale KI · {_settings.Current.AiLocalPreset} · {LocalAiModelCatalog.Get(_settings.Current.AiLocalPreset).DisplayName.Split('–')[1].Trim()}"
        : $"OpenAI-kompatible API · {_settings.Current.AiModel}";
    public string AvailabilityMessage => !IsEnabled ? "KI ist deaktiviert.\nAktiviere und konfiguriere die KI in den Einstellungen."
        : _settings.Current.AiProvider != AiProviderType.LocalLlama ? string.Empty
        : LocalServer.Status == LocalAiStatus.Error ? $"Lokale KI ist nicht verfügbar. {LocalServer.LastError}".Trim()
        : LocalServer.IsReady ? string.Empty : "Lokale KI wird geladen …";
    public event EventHandler? StateChanged;
    private void OnStateChanged(object? sender, EventArgs args) => StateChanged?.Invoke(this, EventArgs.Empty);
    private void OnSettingsChanged() => StateChanged?.Invoke(this, EventArgs.Empty);
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
    public async Task<string> ChatAsync(IReadOnlyList<AiChatRequestMessage> messages, AiRequestOptions options, CancellationToken cancellationToken = default)
    {
        if (!IsEnabled) throw new InvalidOperationException("KI ist deaktiviert.");
        var providerName = _settings.Current.AiProvider.ToString();
        var stopwatch = Stopwatch.StartNew();
        _logger.OperationalInfo($"[AI] Chat request started provider={providerName}");
        try
        {
            var answer = await CreateProvider().ChatAsync(messages, options, cancellationToken);
            answer = Regex.Replace(answer, @"<think>[\s\S]*?</think>", "", RegexOptions.IgnoreCase).Trim();
            if (string.IsNullOrWhiteSpace(answer)) throw new InvalidDataException("Die KI-Antwort ist leer.");
            _logger.OperationalInfo($"[AI] Chat request completed provider={providerName} durationMs={stopwatch.ElapsedMilliseconds}");
            return answer;
        }
        catch (Exception exception)
        {
            _logger.Error($"[AI] Chat request failed provider={providerName} type={exception.GetType().Name} message={exception.Message}");
            throw;
        }
    }
    public async Task InitializeLocalInBackgroundAsync(CancellationToken token = default)
    { if (_settings.Current.AiEnabled && _settings.Current.AiProvider == AiProviderType.LocalLlama) await LocalServer.InstallAndStartAsync(null, token); }
    public void Dispose() { LocalServer.StateChanged -= OnStateChanged; _settings.SettingsChanged -= OnSettingsChanged; LocalServer.Dispose(); _httpClient.Dispose(); }
}

public sealed class LocalLlamaServerManager : IDisposable
{
    public const string Host = "127.0.0.1";
    private static readonly JsonSerializerOptions RuntimeJsonOptions = new() { PropertyNamingPolicy = JsonNamingPolicy.CamelCase, PropertyNameCaseInsensitive = true, WriteIndented = true };
    private readonly SettingsService _settings; private readonly LoggerService _logger; private readonly HttpClient _httpClient; private readonly SemaphoreSlim _gate = new(1, 1);
    private readonly Queue<string> _diagnostics = new(); private Process? _process; private int _port; private string _setupStage = "initialization";
    public LocalLlamaServerManager(SettingsService settings, LoggerService logger, HttpClient httpClient) { _settings = settings; _logger = logger; _httpClient = httpClient; }
    public string RootDirectory { get; } = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Plenaro", "AI");
    public LocalAiComputeMode ActiveComputeMode { get; private set; } = LocalAiComputeMode.Cpu;
    public string? DetectedGpu { get; private set; }
    public bool UsedCpuFallback { get; private set; }
    public string BackendDescription => ActiveComputeMode == LocalAiComputeMode.Gpu ? "Backend: GPU · Vulkan" : "Backend: CPU";
    public string RuntimeDirectory => GetRuntimeDirectory(RootDirectory, ActiveComputeMode);
    public string RuntimeExecutable => Path.Combine(RuntimeDirectory, "llama-server.exe");
    public string RuntimeMetadataPath => Path.Combine(RuntimeDirectory, "runtime.json");
    public string ModelPath(LocalAiPreset preset) => Path.Combine(RootDirectory, "models", preset.ToString().ToLowerInvariant(), LocalAiModelCatalog.Get(preset).FileName);
    public bool IsRunning => _process is { HasExited: false }; public bool IsReady { get; private set; }
    public string? ApiBaseUrl => _port == 0 ? null : $"http://{Host}:{_port}/v1";
    public LocalAiStatus Status { get; private set; } = LocalAiStatus.NotInstalled; public int Progress { get; private set; }
    public string? LastError { get; private set; }
    public event EventHandler? StateChanged;
    internal static string GetRuntimeDirectory(string root, LocalAiComputeMode mode) => Path.Combine(root, "runtime", "llama.cpp", LocalAiRuntimeCatalog.DirectoryName(mode));
    private void SetStatus(LocalAiStatus status, int progress = 0) { Status = status; Progress = progress; StateChanged?.Invoke(this, EventArgs.Empty); }

    public async Task InstallAndStartAsync(IProgress<int>? progress, CancellationToken token = default)
    {
        await _gate.WaitAsync(token);
        try
        {
            if (!_settings.Current.AiEnabled || _settings.Current.AiProvider != AiProviderType.LocalLlama) return;
            LastError = null;
            var preset = _settings.Current.AiLocalPreset;
            ActiveComputeMode = _settings.Current.AiLocalComputeMode; UsedCpuFallback = false; DetectedGpu = null;
            Log($"Local AI setup started preset={preset} backend={LocalAiRuntimeCatalog.Get(ActiveComputeMode).Backend}");
            Stop();
            preset = await EnsureInstalledAsync(progress, token);
            _setupStage = "server-start";
            try { await StartAndWaitForReadyAsync(preset, token); }
            catch (Exception gpuFailure) when (ShouldFallbackToCpu(ActiveComputeMode, gpuFailure))
            {
                Log($"GPU startup failed; falling back to CPU: {gpuFailure.GetType().Name}: {gpuFailure.Message}");
                Stop(); ActiveComputeMode = LocalAiComputeMode.Cpu; UsedCpuFallback = true;
                await EnsureInstalledAsync(progress, token);
                await StartAndWaitForReadyAsync(preset, token);
                LastError = "GPU-Beschleunigung konnte nicht gestartet werden. Lokale KI wird auf CPU gestartet.";
            }
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
            ActiveComputeMode = _settings.Current.AiLocalComputeMode;
            Log($"Local AI setup started preset={_settings.Current.AiLocalPreset} backend={LocalAiRuntimeCatalog.Get(ActiveComputeMode).Backend}");
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
        var runtime = LocalAiRuntimeCatalog.Get(ActiveComputeMode);
        Log($"Required llama.cpp runtime={runtime.Version}");
        var installedMetadata = ReadRuntimeMetadata(RuntimeMetadataPath);
        if (IsRuntimeInstalled(RuntimeExecutable, RuntimeMetadataPath, runtime))
        {
            Log($"Installed llama.cpp runtime={installedMetadata!.Version}");
            Log("Runtime already installed");
        }
        else
        {
            if (installedMetadata != null && (!string.Equals(installedMetadata.Version, runtime.Version, StringComparison.Ordinal) ||
                !string.Equals(installedMetadata.ArchiveSha256, runtime.Sha256, StringComparison.OrdinalIgnoreCase)))
                Log($"Runtime update required installed={installedMetadata.Version} required={runtime.Version}");
            await DownloadRuntimeAsync(progress, token);
        }
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
        var runtime = LocalAiRuntimeCatalog.Get(ActiveComputeMode);
        if (!Uri.TryCreate(runtime.DownloadUrl, UriKind.Absolute, out var url) || url.Scheme != Uri.UriSchemeHttps)
            throw new InvalidOperationException("Die Runtime darf nur über HTTPS geladen werden.");
        SetStatus(LocalAiStatus.DownloadingRuntime);
        _setupStage = "runtime-download";
        Directory.CreateDirectory(RootDirectory); var part = Path.Combine(RootDirectory, $"llama-runtime-{runtime.Backend}.zip.part");
        try
        {
            Log($"Installing llama.cpp runtime={runtime.Version}");
            Log($"Downloading runtime asset={runtime.FileName}");
            await DownloadAsync(url, part, progress, token);
            Log($"Downloaded runtime bytes={new FileInfo(part).Length}");
            var actualSha256 = await CalculateSha256Async(part, token);
            if (!actualSha256.Equals(runtime.Sha256, StringComparison.OrdinalIgnoreCase))
            {
                Log($"Runtime SHA256 mismatch backend={runtime.Backend} asset={runtime.FileName} expected={runtime.Sha256} actual={actualSha256}");
                throw new InvalidDataException("Die heruntergeladene llama.cpp-Runtime konnte nicht verifiziert werden.");
            }
            Log("Runtime SHA256 verified");
            await InstallRuntimeArchiveAsync(part, RuntimeDirectory, runtime, Stop, token);
            Log("Runtime extracted");
            Log($"Runtime installation completed version={runtime.Version}");
        }
        finally { if (File.Exists(part)) File.Delete(part); }
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

    internal static bool IsRuntimeInstalled(string executablePath, string metadataPath, LocalAiRuntimeDefinition required)
    {
        var metadata = ReadRuntimeMetadata(metadataPath);
        return File.Exists(executablePath) && metadata != null &&
            string.Equals(metadata.Provider, "llama.cpp", StringComparison.Ordinal) &&
            string.Equals(metadata.Version, required.Version, StringComparison.Ordinal) &&
            string.Equals(metadata.Backend, required.Backend, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(metadata.Platform, required.Platform, StringComparison.Ordinal) &&
            string.Equals(metadata.ArchiveSha256, required.Sha256, StringComparison.OrdinalIgnoreCase);
    }

    internal static LocalAiRuntimeMetadata? ReadRuntimeMetadata(string metadataPath)
    {
        if (!File.Exists(metadataPath)) return null;
        try
        {
            return JsonSerializer.Deserialize<LocalAiRuntimeMetadata>(File.ReadAllText(metadataPath), RuntimeJsonOptions);
        }
        catch (JsonException) { return null; }
        catch (IOException) { return null; }
        catch (UnauthorizedAccessException) { return null; }
    }

    internal static async Task InstallRuntimeArchiveAsync(string partPath, string runtimeDirectory, LocalAiRuntimeDefinition runtime, Action stopRunningRuntime, CancellationToken token = default)
    {
        var parent = Path.GetDirectoryName(runtimeDirectory) ?? throw new InvalidOperationException("Das Runtime-Verzeichnis besitzt kein übergeordnetes Verzeichnis.");
        Directory.CreateDirectory(parent);
        var staging = Path.Combine(parent, $"llama.cpp.install-{Guid.NewGuid():N}");
        var backup = Path.Combine(parent, $"llama.cpp.backup-{Guid.NewGuid():N}");
        try
        {
            if (!await HasSha256Async(partPath, runtime.Sha256, token))
                throw new InvalidDataException("Die heruntergeladene llama.cpp-Runtime konnte nicht verifiziert werden.");
            ExtractZipSafely(partPath, staging);
            NormalizeRuntimeLayout(staging);
            if (!File.Exists(Path.Combine(staging, "llama-server.exe")))
                throw new InvalidDataException("llama-server.exe fehlt im Runtime-Archiv.");
            var metadata = new LocalAiRuntimeMetadata("llama.cpp", runtime.Version, runtime.Backend, runtime.Platform, runtime.Sha256);
            await File.WriteAllTextAsync(Path.Combine(staging, "runtime.json"), JsonSerializer.Serialize(metadata, RuntimeJsonOptions), token);

            stopRunningRuntime();
            if (Directory.Exists(runtimeDirectory)) Directory.Move(runtimeDirectory, backup);
            try
            {
                Directory.Move(staging, runtimeDirectory);
            }
            catch
            {
                if (Directory.Exists(backup) && !Directory.Exists(runtimeDirectory)) Directory.Move(backup, runtimeDirectory);
                throw;
            }
            if (Directory.Exists(backup)) Directory.Delete(backup, true);
        }
        finally
        {
            if (Directory.Exists(staging)) Directory.Delete(staging, true);
            if (File.Exists(partPath)) File.Delete(partPath);
        }
    }

    private static void NormalizeRuntimeLayout(string runtimeDirectory)
    {
        var runtimeExecutable = Path.Combine(runtimeDirectory, "llama-server.exe");
        if (File.Exists(runtimeExecutable)) return;
        var executable = Directory.EnumerateFiles(runtimeDirectory, "llama-server.exe", SearchOption.AllDirectories).FirstOrDefault();
        if (executable == null) return;
        var source = Path.GetDirectoryName(executable)!;
        foreach (var file in Directory.EnumerateFiles(source)) File.Move(file, Path.Combine(runtimeDirectory, Path.GetFileName(file)), true);
    }
    internal static async Task<string> CalculateSha256Async(string path, CancellationToken token = default) { await using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 131072, true); var hash = await SHA256.HashDataAsync(stream, token); return Convert.ToHexString(hash).ToLowerInvariant(); }
    internal static async Task<bool> HasSha256Async(string path, string expected, CancellationToken token = default) => (await CalculateSha256Async(path, token)).Equals(expected, StringComparison.OrdinalIgnoreCase);

    private async Task StartAndWaitForReadyAsync(LocalAiPreset preset, CancellationToken token)
    {
        _port = GetFreePort(); var info = new ProcessStartInfo(RuntimeExecutable) { UseShellExecute = false, CreateNoWindow = true, RedirectStandardOutput = true, RedirectStandardError = true };
        foreach (var arg in BuildServerArguments(ModelPath(preset), _port, LocalAiModelCatalog.Get(preset).LlamaAlias, ActiveComputeMode)) info.ArgumentList.Add(arg);
        _process = new Process { StartInfo = info, EnableRaisingEvents = true }; _process.OutputDataReceived += CaptureDiagnostic; _process.ErrorDataReceived += CaptureDiagnostic;
        if (!_process.Start()) throw new InvalidOperationException("llama.cpp konnte nicht gestartet werden."); _process.BeginOutputReadLine(); _process.BeginErrorReadLine(); SetStatus(LocalAiStatus.LoadingModel); _logger.Info($"[AI] Local llama.cpp server started host={Host} port={_port}.");
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(token); timeout.CancelAfter(TimeSpan.FromMinutes(3));
        while (!timeout.IsCancellationRequested) { if (_process.HasExited) throw new InvalidOperationException("llama.cpp wurde während des Ladens beendet. " + string.Join(" | ", _diagnostics)); try { using var response = await _httpClient.GetAsync($"http://{Host}:{_port}/health", timeout.Token); if (response.IsSuccessStatusCode) { IsReady = true; SetStatus(LocalAiStatus.Ready, 100); return; } } catch (HttpRequestException) { } await Task.Delay(500, timeout.Token); }
        throw new TimeoutException("llama.cpp wurde nicht innerhalb von drei Minuten bereit.");
    }

    internal static IReadOnlyList<string> BuildServerArguments(string modelPath, int port, string alias, LocalAiComputeMode mode = LocalAiComputeMode.Cpu)
    {
        var arguments = new List<string> { "-m", modelPath, "--host", Host, "--port", port.ToString(), "--ctx-size", "4096", "--parallel", "1", "--alias", alias, "--reasoning", "off" };
        if (mode == LocalAiComputeMode.Gpu) arguments.AddRange(["--n-gpu-layers", "all"]);
        return arguments;
    }
    internal static bool ShouldFallbackToCpu(LocalAiComputeMode mode, Exception exception) => mode == LocalAiComputeMode.Gpu && exception is not OperationCanceledException;
    private void CaptureDiagnostic(object sender, DataReceivedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(e.Data)) return;
        var match = Regex.Match(e.Data, @"(?:Vulkan|device)\s*(?:device)?\s*\d*\s*[:=-]\s*(.+)", RegexOptions.IgnoreCase);
        if (ActiveComputeMode == LocalAiComputeMode.Gpu && match.Success && !string.IsNullOrWhiteSpace(match.Groups[1].Value)) DetectedGpu ??= match.Groups[1].Value.Trim();
        lock (_diagnostics) { _diagnostics.Enqueue(e.Data); while (_diagnostics.Count > 30) _diagnostics.Dequeue(); }
    }
    internal static int GetFreePort() { var listener = new TcpListener(IPAddress.Loopback, 0); listener.Start(); var port = ((IPEndPoint)listener.LocalEndpoint).Port; listener.Stop(); return port; }
    public void Stop() { IsReady = false; if (_process is { HasExited: false }) { _process.Kill(entireProcessTree: true); _process.WaitForExit(5000); } _process?.Dispose(); _process = null; _port = 0; StateChanged?.Invoke(this, EventArgs.Empty); }
    public void Dispose() { Stop(); _gate.Dispose(); }
}
