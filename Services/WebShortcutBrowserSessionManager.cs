using System.Collections.Concurrent;
using System.IO;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.Wpf;
using TaskTool.Models;

namespace TaskTool.Services;

public sealed class WebShortcutBrowserSessionManager : IDisposable
{
    private readonly Dictionary<string, WebShortcutBrowserSession> _sessions = new(StringComparer.Ordinal);
    private readonly SettingsService _settings;
    private readonly LoggerService _logger;
    private Panel? _persistentHost;
    private string? _activeShortcutId;

    public WebShortcutBrowserSessionManager(SettingsService settings, LoggerService logger)
    {
        _settings = settings;
        _logger = logger;
    }

    public static string GetEnvironmentKey(WebShortcutSettings shortcut)
        => $"{shortcut.Id}|cors={(shortcut.DisableWebSecurity ? 1 : 0)}|mixed={(shortcut.AllowInsecureContent ? 1 : 0)}";

    public void AttachPersistentHost(Panel host)
    {
        if (_persistentHost != null && !ReferenceEquals(_persistentHost, host))
            throw new InvalidOperationException("Der persistente Favoriten-Host darf während der Sitzung nicht ausgetauscht werden.");
        _persistentHost = host;
    }

    public WebShortcutBrowserSession Activate(WebShortcutSettings shortcut)
    {
        _activeShortcutId = shortcut.Id;
        var session = GetOrCreate(shortcut);
        if (_persistentHost == null)
            throw new InvalidOperationException("Der persistente Favoriten-Host ist noch nicht verfügbar.");
        _persistentHost.Visibility = Visibility.Visible;
        foreach (var configured in _sessions.Values)
            configured.Browser.Visibility = ReferenceEquals(configured, session) ? Visibility.Visible : Visibility.Collapsed;
        if (session.Browser.Parent == null)
            _persistentHost.Children.Add(session.Browser);
        else if (!ReferenceEquals(session.Browser.Parent, _persistentHost))
            throw new InvalidOperationException("Eine Favoriten-Browserinstanz darf nicht zwischen Visual Parents verschoben werden.");
        return session;
    }

    public void HideAll()
    {
        _activeShortcutId = null;
        if (_persistentHost != null)
            _persistentHost.Visibility = Visibility.Collapsed;
        foreach (var session in _sessions.Values)
            session.Browser.Visibility = Visibility.Collapsed;
    }

    public WebShortcutBrowserSession GetOrCreate(WebShortcutSettings shortcut)
    {
        var key = GetEnvironmentKey(shortcut);
        if (_sessions.TryGetValue(shortcut.Id, out var existing) && !string.Equals(existing.EnvironmentKey, key, StringComparison.Ordinal))
        {
            existing.Dispose();
            _sessions.Remove(shortcut.Id);
            existing = null;
        }

        if (existing == null)
        {
            existing = new WebShortcutBrowserSession(shortcut, _settings, _logger);
            _sessions.Add(shortcut.Id, existing);
        }
        else
        {
            existing.UpdateConfiguration(shortcut);
        }

        return existing;
    }

    public void Synchronize(IReadOnlyCollection<WebShortcutSettings> shortcuts)
    {
        var configured = shortcuts.ToDictionary(x => x.Id, StringComparer.Ordinal);
        foreach (var entry in _sessions.ToList())
        {
            if (!configured.TryGetValue(entry.Key, out var shortcut)
                || !string.Equals(entry.Value.EnvironmentKey, GetEnvironmentKey(shortcut), StringComparison.Ordinal))
            {
                entry.Value.Dispose();
                _sessions.Remove(entry.Key);
                continue;
            }
            entry.Value.UpdateConfiguration(shortcut);
        }
        if (_activeShortcutId != null && configured.TryGetValue(_activeShortcutId, out var active))
        {
            var session = Activate(active);
            _ = InitializeActivatedSessionAsync(session);
        }
    }

    private async Task InitializeActivatedSessionAsync(WebShortcutBrowserSession session)
    {
        try { await session.EnsureInitializedAsync(); }
        catch (Exception ex)
        {
            _logger.Warning($"[WebShortcutBrowser] shortcutId={session.ShortcutId} action=reinitialize-failed errorType={ex.GetType().Name}");
        }
    }

    public void Dispose()
    {
        foreach (var session in _sessions.Values) session.Dispose();
        _sessions.Clear();
    }
}

public sealed class WebShortcutBrowserSession : IDisposable
{
    private readonly SettingsService _settingsService;
    private readonly LoggerService _logger;
    private readonly ConcurrentDictionary<string, CoreWebView2WebResourceContext> _requestContexts = new(StringComparer.Ordinal);
    private readonly ConcurrentDictionary<string, DevToolsRequest> _devToolsRequests = new(StringComparer.Ordinal);
    private CancellationTokenSource _lifetimeCts = new();
    private CancellationTokenSource? _loginCts;
    private Task? _initializationTask;
    private WebShortcutSettings _settings;
    private bool _disposed;
    private CoreWebView2DevToolsProtocolEventReceiver? _requestWillBeSentReceiver;
    private CoreWebView2DevToolsProtocolEventReceiver? _loadingFailedReceiver;
    private CoreWebView2DevToolsProtocolEventReceiver? _loadingFinishedReceiver;

    public string ShortcutId => _settings.Id;
    public string EnvironmentKey { get; }
    public string SecurityMode => GetProfileName(_settings);
    public WebView2 Browser { get; } = new();
    public CoreWebView2Environment? Environment { get; private set; }
    public bool Initialized { get; private set; }
    public string? CurrentUrl => Browser.Source?.ToString();
    public string LastStatus { get; private set; } = string.Empty;
    public event Action<string>? StatusChanged;

    public WebShortcutBrowserSession(WebShortcutSettings settings, SettingsService settingsService, LoggerService logger)
    {
        _settings = settings;
        _settingsService = settingsService;
        _logger = logger;
        EnvironmentKey = WebShortcutBrowserSessionManager.GetEnvironmentKey(settings);
    }

    public async Task EnsureInitializedAsync()
    {
        _initializationTask ??= InitializeAsync();
        try { await _initializationTask; }
        catch
        {
            _initializationTask = null;
            throw;
        }
    }

    public void UpdateConfiguration(WebShortcutSettings settings)
    {
        if (_disposed) return;
        var urlChanged = !string.Equals(_settings.Url, settings.Url, StringComparison.Ordinal);
        _settings = settings;
        if (urlChanged && Initialized && TryGetHttpUri(settings.Url, out var uri))
        {
            _logger.Info($"[WebShortcutBrowser] shortcutId={ShortcutId} action=navigate-configured-url host={uri.Host}");
            Browser.CoreWebView2.Navigate(uri.ToString());
        }
    }

    private async Task InitializeAsync()
    {
        SetStatus("Webseite wird geladen …");
        try
        {
            var profile = GetProfileName(_settings);
            var userDataFolder = Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData),
                "Plenaro", "WebView2", "Shortcuts", ShortcutId, profile);
            var arguments = GetBrowserArguments(_settings);
            Environment = string.IsNullOrEmpty(arguments)
                ? await CoreWebView2Environment.CreateAsync(null, userDataFolder)
                : await CoreWebView2Environment.CreateAsync(null, userDataFolder,
                    new CoreWebView2EnvironmentOptions { AdditionalBrowserArguments = arguments });
            _lifetimeCts.Token.ThrowIfCancellationRequested();
            await Browser.EnsureCoreWebView2Async(Environment);
            _lifetimeCts.Token.ThrowIfCancellationRequested();
            await ConfigureBrowserEventsAsync();
            Initialized = true;
            SetStatus(string.Empty);
            _logger.Info($"[WebShortcutBrowser] shortcutId={ShortcutId} environmentKey={EnvironmentKey} action=create");
            if (TryGetHttpUri(_settings.Url, out var uri))
            {
                _logger.Info($"[WebShortcutBrowser] shortcutId={ShortcutId} action=navigate host={uri.Host}");
                Browser.CoreWebView2.Navigate(uri.ToString());
            }
        }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex)
        {
            if (!_disposed) SetStatus($"Webseite konnte nicht geöffnet werden: {ex.Message}");
            throw;
        }
    }

    private async Task ConfigureBrowserEventsAsync()
    {
        var core = Browser.CoreWebView2;
        core.NavigationStarting += NavigationStarting;
        core.NavigationCompleted += NavigationCompleted;
        core.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.All);
        core.WebResourceRequested += WebResourceRequested;
        core.WebResourceResponseReceived += WebResourceResponseReceived;
        _requestWillBeSentReceiver = core.GetDevToolsProtocolEventReceiver("Network.requestWillBeSent");
        _loadingFailedReceiver = core.GetDevToolsProtocolEventReceiver("Network.loadingFailed");
        _loadingFinishedReceiver = core.GetDevToolsProtocolEventReceiver("Network.loadingFinished");
        _requestWillBeSentReceiver.DevToolsProtocolEventReceived += DevToolsRequestWillBeSent;
        _loadingFailedReceiver.DevToolsProtocolEventReceived += DevToolsLoadingFailed;
        _loadingFinishedReceiver.DevToolsProtocolEventReceived += DevToolsLoadingFinished;
        await core.CallDevToolsProtocolMethodAsync("Network.enable", "{}");
    }

    private void NavigationStarting(object? sender, CoreWebView2NavigationStartingEventArgs e)
    {
        _loginCts?.Cancel();
        if (!TryGetHttpUri(e.Uri, out _))
        {
            e.Cancel = true;
            SetStatus("Navigation blockiert.");
        }
    }

    private void NavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
    {
        if (!e.IsSuccess)
        {
            var host = TryGetHttpUri(CurrentUrl, out var failedUri) ? failedUri.Host : "unknown";
            _logger.Warning($"[WebShortcutRequest] shortcutId={ShortcutId} host={host} status=0 result=network-failure webError={e.WebErrorStatus} utc={DateTime.UtcNow:O}");
            SetStatus("API-Aufruf fehlgeschlagen – Details siehe Log.");
            return;
        }

        if (!TryGetHttpUri(CurrentUrl, out var current)) return;
        _logger.Info($"[WebShortcut] shortcutId={ShortcutId} host={current.Host} navigation=success");
        if (!_settings.AutoLogin || !Trusted(current)) return;
        var password = _settingsService.GetWebShortcutPassword(_settings);
        if (password.Length == 0 || _settings.Username.Length == 0) return;
        _loginCts?.Cancel();
        _loginCts?.Dispose();
        _loginCts = CancellationTokenSource.CreateLinkedTokenSource(_lifetimeCts.Token);
        _ = LoginAsync(_settings.Username, password, current.Host, _loginCts.Token);
    }

    private void WebResourceRequested(object? sender, CoreWebView2WebResourceRequestedEventArgs e)
    {
        if (e.ResourceContext is not (CoreWebView2WebResourceContext.Fetch or CoreWebView2WebResourceContext.XmlHttpRequest)) return;
        _requestContexts[RequestKey(e.Request)] = e.ResourceContext;
    }

    private void WebResourceResponseReceived(object? sender, CoreWebView2WebResourceResponseReceivedEventArgs e)
    {
        var status = e.Response.StatusCode;
        var key = RequestKey(e.Request);
        _requestContexts.TryRemove(key, out var context);
        if (status < 400 || !TryGetHttpUri(e.Request.Uri, out var uri)) return;
        var result = status == 403 && string.Equals(e.Request.Method, "OPTIONS", StringComparison.OrdinalIgnoreCase)
            ? "server-rejected-preflight"
            : "http-error";
        _logger.Warning($"[WebShortcutRequest] shortcutId={ShortcutId} host={uri.Host} method={e.Request.Method} context={context} status={status} result={result} utc={DateTime.UtcNow:O}");
        if (context is CoreWebView2WebResourceContext.Fetch or CoreWebView2WebResourceContext.XmlHttpRequest)
            SetStatus("API-Aufruf fehlgeschlagen – Details siehe Log.");
    }

    private void DevToolsRequestWillBeSent(object? sender, CoreWebView2DevToolsProtocolEventReceivedEventArgs e)
    {
        try
        {
            using var document = JsonDocument.Parse(e.ParameterObjectAsJson);
            var root = document.RootElement;
            if (!root.TryGetProperty("requestId", out var requestIdElement)
                || !root.TryGetProperty("request", out var request)
                || !request.TryGetProperty("url", out var urlElement)) return;
            var requestId = requestIdElement.GetString();
            var url = urlElement.GetString();
            if (string.IsNullOrEmpty(requestId) || !TryGetHttpUri(url, out var uri)) return;
            var method = request.TryGetProperty("method", out var methodElement) ? methodElement.GetString() ?? "unknown" : "unknown";
            var context = root.TryGetProperty("type", out var typeElement) ? typeElement.GetString() ?? "Other" : "Other";
            _devToolsRequests[requestId] = new DevToolsRequest(uri.Host, method, context);
        }
        catch (JsonException) { }
    }

    private void DevToolsLoadingFailed(object? sender, CoreWebView2DevToolsProtocolEventReceivedEventArgs e)
    {
        try
        {
            using var document = JsonDocument.Parse(e.ParameterObjectAsJson);
            var root = document.RootElement;
            var requestId = root.TryGetProperty("requestId", out var requestIdElement) ? requestIdElement.GetString() : null;
            if (string.IsNullOrEmpty(requestId) || !_devToolsRequests.TryRemove(requestId, out var request)) return;
            if (request.Context is not ("Fetch" or "XHR" or "EventSource")) return;
            var blockedReason = root.TryGetProperty("blockedReason", out var blockedElement) ? blockedElement.GetString() ?? string.Empty : string.Empty;
            var errorText = root.TryGetProperty("errorText", out var errorElement) ? errorElement.GetString() ?? string.Empty : string.Empty;
            var hasCorsStatus = root.TryGetProperty("corsErrorStatus", out _);
            var result = hasCorsStatus ? "browser-cors"
                : blockedReason.Contains("mixed", StringComparison.OrdinalIgnoreCase) ? "mixed-content"
                : errorText.Contains("CERT", StringComparison.OrdinalIgnoreCase) ? "tls-error"
                : "network-failure";
            _logger.Warning($"[WebShortcutRequest] shortcutId={ShortcutId} host={request.Host} method={request.Method} context={request.Context} status=0 result={result} utc={DateTime.UtcNow:O}");
            SetStatus("API-Aufruf fehlgeschlagen – Details siehe Log.");
        }
        catch (JsonException) { }
    }

    private void DevToolsLoadingFinished(object? sender, CoreWebView2DevToolsProtocolEventReceivedEventArgs e)
    {
        try
        {
            using var document = JsonDocument.Parse(e.ParameterObjectAsJson);
            if (document.RootElement.TryGetProperty("requestId", out var requestIdElement)
                && requestIdElement.GetString() is { Length: > 0 } requestId) _devToolsRequests.TryRemove(requestId, out _);
        }
        catch (JsonException) { }
    }

    private bool Trusted(Uri current)
        => TryGetHttpUri(_settings.Url, out var configured)
           && current.Scheme == Uri.UriSchemeHttps
           && configured.Scheme == Uri.UriSchemeHttps
           && current.Scheme.Equals(configured.Scheme, StringComparison.OrdinalIgnoreCase)
           && current.Host.Equals(configured.Host, StringComparison.OrdinalIgnoreCase)
           && current.Port == configured.Port;

    private async Task LoginAsync(string username, string password, string host, CancellationToken token)
    {
        var schedule = new[] { 0, 300, 800, 1500 };
        var prior = 0;
        foreach (var at in schedule)
        {
            try
            {
                if (at > prior) await Task.Delay(at - prior, token);
                prior = at;
                var result = await FillAsync(username, password);
                token.ThrowIfCancellationRequested();
                _logger.Info($"[WebShortcutLogin] shortcutId={ShortcutId} host={host} loginPageDetected={result.Detected.ToString().ToLowerInvariant()} usernameFieldFound={result.User.ToString().ToLowerInvariant()} passwordFieldFound={result.Password.ToString().ToLowerInvariant()} submitted={result.Submitted.ToString().ToLowerInvariant()}");
                if (result.Detected) return;
            }
            catch (OperationCanceledException) { return; }
            catch (Exception ex)
            {
                if (!token.IsCancellationRequested) _logger.Warning($"[WebShortcutLogin] shortcutId={ShortcutId} host={host} errorType={ex.GetType().Name}");
                return;
            }
        }
    }

    private async Task<LoginResult> FillAsync(string username, string password)
    {
        var userJson = JsonSerializer.Serialize(username);
        var passwordJson = JsonSerializer.Serialize(password);
        var script = $$"""(()=>{const q=s=>s.map(x=>document.querySelector(x)).find(Boolean)||null;const u=q(['input[autocomplete="username"]','input[type="email"]','input[name="username"]','input[name="user"]','input[name="User"]','input[name="login"]','input[id*="user" i]','input[id*="login" i]']);const p=q(['input[type="password"]','input[name="password"]','input[name="Password"]','input[id*="password" i]']);const f=p?.closest('form');const b=f&&['button[type="submit"]','input[type="submit"]','button[name*="login" i]','button[id*="login" i]'].map(x=>f.querySelector(x)).find(Boolean);const r={detected:!!(u&&p),user:!!u,password:!!p,submitted:false};if(!u||!p)return r;const set=(e,v)=>{const s=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value')?.set;s?s.call(e,v):e.value=v;e.dispatchEvent(new Event('input',{bubbles:true}));e.dispatchEvent(new Event('change',{bubbles:true}));};set(u,{{userJson}});set(p,{{passwordJson}});if(b){b.click();r.submitted=true;}else if(f?.requestSubmit){f.requestSubmit();r.submitted=true;}return r;})()""";
        using var document = JsonDocument.Parse(await Browser.CoreWebView2.ExecuteScriptAsync(script));
        var result = document.RootElement;
        return new(result.GetProperty("detected").GetBoolean(), result.GetProperty("user").GetBoolean(),
            result.GetProperty("password").GetBoolean(), result.GetProperty("submitted").GetBoolean());
    }

    private static string RequestKey(CoreWebView2WebResourceRequest request) => $"{request.Method}\n{request.Uri}";
    private static bool TryGetHttpUri(string? value, out Uri uri)
        => Uri.TryCreate(value, UriKind.Absolute, out uri!) && (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps);
    private static string GetProfileName(WebShortcutSettings settings) => (settings.DisableWebSecurity, settings.AllowInsecureContent) switch
    {
        (false, false) => "normal",
        (true, false) => "cors",
        (false, true) => "mixed",
        (true, true) => "cors-mixed"
    };
    private static string GetBrowserArguments(WebShortcutSettings settings)
        => string.Join(" ", new[]
        {
            settings.DisableWebSecurity ? "--disable-web-security" : string.Empty,
            settings.AllowInsecureContent ? "--allow-running-insecure-content" : string.Empty
        }.Where(x => x.Length > 0));

    private void SetStatus(string status)
    {
        LastStatus = status;
        StatusChanged?.Invoke(status);
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _lifetimeCts.Cancel();
        _loginCts?.Cancel();
        if (Browser.CoreWebView2 != null)
        {
            Browser.CoreWebView2.NavigationStarting -= NavigationStarting;
            Browser.CoreWebView2.NavigationCompleted -= NavigationCompleted;
            Browser.CoreWebView2.WebResourceRequested -= WebResourceRequested;
            Browser.CoreWebView2.WebResourceResponseReceived -= WebResourceResponseReceived;
        }
        if (_requestWillBeSentReceiver != null) _requestWillBeSentReceiver.DevToolsProtocolEventReceived -= DevToolsRequestWillBeSent;
        if (_loadingFailedReceiver != null) _loadingFailedReceiver.DevToolsProtocolEventReceived -= DevToolsLoadingFailed;
        if (_loadingFinishedReceiver != null) _loadingFinishedReceiver.DevToolsProtocolEventReceived -= DevToolsLoadingFinished;
        if (Browser.Parent is Panel parent) parent.Children.Remove(Browser);
        Browser.Dispose();
        _loginCts?.Dispose();
        _lifetimeCts.Dispose();
    }

    private sealed record LoginResult(bool Detected, bool User, bool Password, bool Submitted);
    private sealed record DevToolsRequest(string Host, string Method, string Context);
}
