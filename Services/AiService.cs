using System.Diagnostics;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Text.Json;
using TaskTool.Models;

namespace TaskTool.Services;

public interface IAiService
{
    Task<string> SendAsync(string systemPrompt, string userPrompt, CancellationToken cancellationToken = default);
}

public sealed class OpenAiCompatibleAiProvider : IAiService
{
    private readonly HttpClient _httpClient;
    private readonly Uri _endpoint;
    private readonly string _model;
    private readonly string _apiKey;

    public OpenAiCompatibleAiProvider(HttpClient httpClient, string baseUrl, string model, string apiKey = "")
    {
        _httpClient = httpClient;
        _endpoint = BuildEndpoint(baseUrl);
        _model = string.IsNullOrWhiteSpace(model) ? throw new ArgumentException("Ein Modellname ist erforderlich.", nameof(model)) : model.Trim();
        _apiKey = apiKey;
    }

    public async Task<string> SendAsync(string systemPrompt, string userPrompt, CancellationToken cancellationToken = default)
    {
        using var request = new HttpRequestMessage(HttpMethod.Post, _endpoint)
        {
            Content = JsonContent.Create(new
            {
                model = _model,
                messages = new[] { new { role = "system", content = systemPrompt }, new { role = "user", content = userPrompt } },
                temperature = 0,
                max_tokens = 32
            })
        };
        if (!string.IsNullOrWhiteSpace(_apiKey)) request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
        using var response = await _httpClient.SendAsync(request, cancellationToken);
        response.EnsureSuccessStatusCode();
        await using var stream = await response.Content.ReadAsStreamAsync(cancellationToken);
        using var document = await JsonDocument.ParseAsync(stream, cancellationToken: cancellationToken);
        if (!document.RootElement.TryGetProperty("choices", out var choices) || choices.GetArrayLength() == 0
            || !choices[0].TryGetProperty("message", out var message)
            || !message.TryGetProperty("content", out var content))
            throw new InvalidDataException("Die KI-Antwort enthält keinen Text im erwarteten Chat-Completions-Format.");
        return content.GetString()?.Trim() ?? string.Empty;
    }

    internal static Uri BuildEndpoint(string baseUrl)
    {
        if (!Uri.TryCreate(baseUrl?.Trim(), UriKind.Absolute, out var uri) || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
            throw new UriFormatException("Die API Basis-URL muss eine absolute HTTP- oder HTTPS-URL sein.");
        var root = uri.ToString().TrimEnd('/');
        return new Uri(root.EndsWith("/v1", StringComparison.OrdinalIgnoreCase)
            ? root + "/chat/completions"
            : root + "/v1/chat/completions");
    }
}

public sealed class AiService : IDisposable
{
    public const string TestSystemPrompt = "Folge der Benutzeranweisung exakt. Gib keine zusätzlichen Erklärungen aus.";
    public const string TestUserPrompt = "Antworte ausschließlich mit exakt: Test erfolgreich";
    private readonly SettingsService _settings;
    private readonly LocalLlamaServerManager _localServer;
    private readonly HttpClient _httpClient;

    public AiService(SettingsService settings, LoggerService logger, HttpClient? httpClient = null)
    {
        _settings = settings;
        _httpClient = httpClient ?? new HttpClient { Timeout = TimeSpan.FromSeconds(60) };
        _localServer = new LocalLlamaServerManager(settings, logger, _httpClient);
    }

    public LocalLlamaServerManager LocalServer => _localServer;

    public IAiService CreateProvider()
    {
        var config = _settings.Current;
        return config.AiProvider switch
        {
            AiProviderType.OpenAiCompatible => new OpenAiCompatibleAiProvider(_httpClient, config.AiApiBaseUrl, config.AiModel, _settings.GetAiApiKey()),
            AiProviderType.LocalLlama => new OpenAiCompatibleAiProvider(_httpClient, $"http://127.0.0.1:{config.AiLocalServerPort}/v1", config.AiModel),
            _ => throw new InvalidOperationException("Der gewählte KI-Provider wird nicht unterstützt.")
        };
    }

    public async Task<string> TestAsync(CancellationToken cancellationToken = default)
    {
        if (!_settings.Current.AiEnabled) throw new InvalidOperationException("KI ist deaktiviert.");
        return await CreateProvider().SendAsync(TestSystemPrompt, TestUserPrompt, cancellationToken);
    }

    public void Dispose() { _localServer.Dispose(); _httpClient.Dispose(); }
}

public sealed class LocalLlamaServerManager : IDisposable
{
    private readonly SettingsService _settings;
    private readonly LoggerService _logger;
    private readonly HttpClient _httpClient;
    private Process? _process;

    public LocalLlamaServerManager(SettingsService settings, LoggerService logger, HttpClient httpClient)
    { _settings = settings; _logger = logger; _httpClient = httpClient; }

    public bool IsRunning => _process is { HasExited: false };

    public void Start()
    {
        if (!_settings.Current.AiEnabled || _settings.Current.AiProvider != AiProviderType.LocalLlama)
            throw new InvalidOperationException("Lokale KI ist nicht aktiviert.");
        if (IsRunning) return;
        var executable = _settings.Current.AiLocalServerExecutablePath;
        var model = _settings.Current.AiLocalModelPath;
        if (!File.Exists(executable)) throw new FileNotFoundException("Die llama.cpp-Serverdatei wurde nicht gefunden.");
        if (!File.Exists(model)) throw new FileNotFoundException("Die lokale Modelldatei wurde nicht gefunden.");
        var startInfo = new ProcessStartInfo(executable) { UseShellExecute = false, CreateNoWindow = true };
        startInfo.ArgumentList.Add("-m"); startInfo.ArgumentList.Add(model);
        startInfo.ArgumentList.Add("--host"); startInfo.ArgumentList.Add("127.0.0.1");
        startInfo.ArgumentList.Add("--port"); startInfo.ArgumentList.Add(_settings.Current.AiLocalServerPort.ToString());
        _process = Process.Start(startInfo) ?? throw new InvalidOperationException("llama.cpp konnte nicht gestartet werden.");
        _logger.Info($"[AI] Local llama.cpp server started port={_settings.Current.AiLocalServerPort}.");
    }

    public void Stop()
    {
        if (!IsRunning) { _process?.Dispose(); _process = null; return; }
        _process!.Kill(entireProcessTree: true); _process.WaitForExit(5000); _process.Dispose(); _process = null;
        _logger.Info("[AI] Local llama.cpp server stopped.");
    }

    public async Task DownloadModelAsync(IProgress<int>? progress = null, CancellationToken cancellationToken = default)
    {
        var url = _settings.Current.AiLocalModelDownloadUrl;
        var destination = _settings.Current.AiLocalModelPath;
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) || uri.Scheme is not ("http" or "https")) throw new UriFormatException("Die Modell-URL ist ungültig.");
        if (string.IsNullOrWhiteSpace(destination)) throw new InvalidOperationException("Bitte einen Zielpfad für das Modell angeben.");
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(destination))!);
        var temporary = destination + ".download";
        try
        {
            using var response = await _httpClient.GetAsync(uri, HttpCompletionOption.ResponseHeadersRead, cancellationToken);
            response.EnsureSuccessStatusCode();
            var total = response.Content.Headers.ContentLength;
            await using var input = await response.Content.ReadAsStreamAsync(cancellationToken);
            await using var output = new FileStream(temporary, FileMode.Create, FileAccess.Write, FileShare.None, 81920, true);
            var buffer = new byte[81920]; long received = 0; int read;
            while ((read = await input.ReadAsync(buffer, cancellationToken)) > 0)
            { await output.WriteAsync(buffer.AsMemory(0, read), cancellationToken); received += read; if (total > 0) progress?.Report((int)(received * 100 / total)); }
            File.Move(temporary, destination, true); progress?.Report(100);
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }

    public void Dispose() => Stop();
}
