using System.Collections.Concurrent;
using System.IO;
using System.Text.Json;
using System.Windows.Controls;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.Wpf;
using TaskTool.Services;
using TaskTool.ViewModels;

namespace TaskTool.Views;

public partial class WebShortcutView : UserControl
{
    private static readonly ConcurrentDictionary<string, Task<CoreWebView2Environment>> EnvironmentCache = new(StringComparer.Ordinal);
    private WebView2? _browser;
    private CancellationTokenSource? _switchCts;
    private CancellationTokenSource? _loginCts;
    private long _navigationGeneration;
    private string? _activeEnvironmentKey;
    private string? _activeShortcutId;
    private WebShortcutViewModel? _subscribedViewModel;

    public WebShortcutView()
    {
        InitializeComponent();
        DataContextChanged += OnDataContextChanged;
    }

    private void OnDataContextChanged(object sender, System.Windows.DependencyPropertyChangedEventArgs e)
    {
        if (_subscribedViewModel != null) _subscribedViewModel.BrowserConfigurationChanged -= OnBrowserConfigurationChanged;
        _subscribedViewModel = e.NewValue as WebShortcutViewModel;
        if (_subscribedViewModel != null) _subscribedViewModel.BrowserConfigurationChanged += OnBrowserConfigurationChanged;
        _ = SwitchAsync();
    }

    private void OnBrowserConfigurationChanged() => _ = SwitchAsync();

    private async Task SwitchAsync()
    {
        _switchCts?.Cancel();
        _switchCts?.Dispose();
        _loginCts?.Cancel();
        _loginCts?.Dispose();
        var switchCts = _switchCts = new CancellationTokenSource();
        var token = switchCts.Token;
        var generation = ++_navigationGeneration;

        if (DataContext is not WebShortcutViewModel viewModel
            || !Uri.TryCreate(viewModel.Url, UriKind.Absolute, out var uri)
            || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
        {
            TearDownBrowser();
            BrowserStatus.Text = "Webseite besitzt keine gültige http/https URL.";
            return;
        }

        var shortcutId = viewModel.ShortcutId;
        var environmentKey = viewModel.EnvironmentKey;
        BrowserStatus.Text = "Webseite wird geladen …";

        try
        {
            var browser = _browser;
            if (browser?.CoreWebView2 == null || !string.Equals(_activeEnvironmentKey, environmentKey, StringComparison.Ordinal))
            {
                TearDownBrowser();
                browser = new WebView2();
                _browser = browser;
                _activeEnvironmentKey = environmentKey;
                _activeShortcutId = shortcutId;
                BrowserHost.Children.Add(browser);
                ServiceLocator.Logger.Info($"[WebShortcutBrowser] shortcutId={shortcutId} environmentKey={environmentKey} action=create");

                var environment = await GetEnvironmentAsync(viewModel.Shortcut, environmentKey);
                if (!IsCurrent(generation, token, shortcutId, browser))
                {
                    LogStale(shortcutId);
                    return;
                }

                await browser.EnsureCoreWebView2Async(environment);
                if (!IsCurrent(generation, token, shortcutId, browser))
                {
                    LogStale(shortcutId);
                    return;
                }

                browser.CoreWebView2.NavigationStarting += Starting;
                browser.CoreWebView2.NavigationCompleted += Completed;
            }

            if (!IsCurrent(generation, token, shortcutId, browser))
            {
                LogStale(shortcutId);
                return;
            }

            BrowserStatus.Text = string.Empty;
            ServiceLocator.Logger.Info($"[WebShortcutBrowser] shortcutId={shortcutId} action=navigate host={uri.Host}");
            browser.CoreWebView2.Navigate(uri.ToString());
        }
        catch (Exception ex)
        {
            if (!IsCurrent(generation, token, shortcutId, _browser))
            {
                LogStale(shortcutId);
                return;
            }
            BrowserStatus.Text = $"Webseite konnte nicht geöffnet werden: {ex.Message}";
        }
    }

    private static async Task<CoreWebView2Environment> GetEnvironmentAsync(TaskTool.Models.WebShortcutSettings shortcut, string environmentKey)
    {
        var task = EnvironmentCache.GetOrAdd(environmentKey, _ => CreateEnvironmentAsync(shortcut));
        try { return await task; }
        catch
        {
            EnvironmentCache.TryRemove(environmentKey, out _);
            throw;
        }
    }

    private static Task<CoreWebView2Environment> CreateEnvironmentAsync(TaskTool.Models.WebShortcutSettings shortcut)
    {
        var securityMode = shortcut.DisableWebSecurity ? "cors-disabled" : "normal";
        var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "Plenaro", "WebView2", "Shortcuts", shortcut.Id, securityMode);
        if (!shortcut.DisableWebSecurity) return CoreWebView2Environment.CreateAsync(null, path);
        var options = new CoreWebView2EnvironmentOptions { AdditionalBrowserArguments = "--disable-web-security" };
        return CoreWebView2Environment.CreateAsync(null, path, options);
    }

    private bool IsCurrent(long generation, CancellationToken token, string shortcutId, WebView2? browser)
        => !token.IsCancellationRequested
           && generation == _navigationGeneration
           && ReferenceEquals(browser, _browser)
           && DataContext is WebShortcutViewModel current
           && string.Equals(current.ShortcutId, shortcutId, StringComparison.Ordinal);

    private static void LogStale(string shortcutId)
        => ServiceLocator.Logger.Info($"[WebShortcutBrowser] shortcutId={shortcutId} action=cancel-stale-navigation");

    private void TearDownBrowser()
    {
        if (_browser?.CoreWebView2 != null)
        {
            _browser.CoreWebView2.NavigationStarting -= Starting;
            _browser.CoreWebView2.NavigationCompleted -= Completed;
        }
        if (_browser != null)
        {
            BrowserHost.Children.Remove(_browser);
            _browser.Dispose();
        }
        _browser = null;
        _activeEnvironmentKey = null;
        _activeShortcutId = null;
    }

    private void Starting(object? sender, CoreWebView2NavigationStartingEventArgs e)
    {
        if (!IsActiveSender(sender, out _)) return;
        _loginCts?.Cancel();
        if (!Uri.TryCreate(e.Uri, UriKind.Absolute, out var uri)
            || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
        {
            e.Cancel = true;
            BrowserStatus.Text = "Navigation blockiert.";
        }
    }

    private void Completed(object? sender, CoreWebView2NavigationCompletedEventArgs e)
    {
        if (!e.IsSuccess || !IsActiveSender(sender, out var viewModel) || _browser == null) return;
        if (!Uri.TryCreate(_browser.Source?.ToString(), UriKind.Absolute, out var current)) return;
        ServiceLocator.Logger.Info($"[WebShortcut] shortcutId={viewModel.ShortcutId} host={current.Host} navigation=success");
        if (!viewModel.Shortcut.AutoLogin || !Trusted(viewModel, current)) return;
        var password = ServiceLocator.Settings.GetWebShortcutPassword(viewModel.Shortcut);
        if (password.Length == 0 || viewModel.Shortcut.Username.Length == 0) return;
        _loginCts?.Cancel();
        _loginCts?.Dispose();
        var loginCts = _loginCts = new CancellationTokenSource();
        _ = LoginAsync(viewModel, _browser, password, current.Host, _navigationGeneration, loginCts.Token);
    }

    private bool IsActiveSender(object? sender, out WebShortcutViewModel viewModel)
    {
        viewModel = DataContext as WebShortcutViewModel ?? null!;
        return viewModel != null
               && ReferenceEquals(sender, _browser?.CoreWebView2)
               && string.Equals(viewModel.ShortcutId, _activeShortcutId, StringComparison.Ordinal);
    }

    private static bool Trusted(WebShortcutViewModel viewModel, Uri current)
        => Uri.TryCreate(viewModel.Url, UriKind.Absolute, out var configured)
           && current.Scheme == Uri.UriSchemeHttps
           && configured.Scheme == Uri.UriSchemeHttps
           && current.Scheme.Equals(configured.Scheme, StringComparison.OrdinalIgnoreCase)
           && current.Host.Equals(configured.Host, StringComparison.OrdinalIgnoreCase)
           && current.Port == configured.Port;

    private async Task LoginAsync(WebShortcutViewModel viewModel, WebView2 browser, string password, string host, long generation, CancellationToken token)
    {
        var schedule = new[] { 0, 300, 800, 1500 };
        var prior = 0;
        foreach (var at in schedule)
        {
            try
            {
                if (at > prior) await Task.Delay(at - prior, token);
                prior = at;
                if (!IsCurrent(generation, token, viewModel.ShortcutId, browser)) return;
                var result = await FillAsync(browser, viewModel.Shortcut.Username, password);
                if (!IsCurrent(generation, token, viewModel.ShortcutId, browser)) return;
                ServiceLocator.Logger.Info($"[WebShortcutLogin] shortcutId={viewModel.ShortcutId} host={host} loginPageDetected={result.Detected.ToString().ToLowerInvariant()} usernameFieldFound={result.User.ToString().ToLowerInvariant()} passwordFieldFound={result.Password.ToString().ToLowerInvariant()} submitted={result.Submitted.ToString().ToLowerInvariant()}");
                if (result.Detected) return;
            }
            catch (OperationCanceledException) { return; }
            catch (Exception ex)
            {
                if (IsCurrent(generation, token, viewModel.ShortcutId, browser))
                    ServiceLocator.Logger.Warning($"[WebShortcutLogin] shortcutId={viewModel.ShortcutId} host={host} errorType={ex.GetType().Name}");
                return;
            }
        }
    }

    private static async Task<LoginResult> FillAsync(WebView2 browser, string username, string password)
    {
        var userJson = JsonSerializer.Serialize(username);
        var passwordJson = JsonSerializer.Serialize(password);
        var script = $$"""(()=>{const q=s=>s.map(x=>document.querySelector(x)).find(Boolean)||null;const u=q(['input[autocomplete="username"]','input[type="email"]','input[name="username"]','input[name="user"]','input[name="User"]','input[name="login"]','input[id*="user" i]','input[id*="login" i]']);const p=q(['input[type="password"]','input[name="password"]','input[name="Password"]','input[id*="password" i]']);const f=p?.closest('form');const b=f&&['button[type="submit"]','input[type="submit"]','button[name*="login" i]','button[id*="login" i]'].map(x=>f.querySelector(x)).find(Boolean);const r={detected:!!(u&&p),user:!!u,password:!!p,submitted:false};if(!u||!p)return r;const set=(e,v)=>{const s=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value')?.set;s?s.call(e,v):e.value=v;e.dispatchEvent(new Event('input',{bubbles:true}));e.dispatchEvent(new Event('change',{bubbles:true}));};set(u,{{userJson}});set(p,{{passwordJson}});if(b){b.click();r.submitted=true;}else if(f?.requestSubmit){f.requestSubmit();r.submitted=true;}return r;})()""";
        using var document = JsonDocument.Parse(await browser.CoreWebView2.ExecuteScriptAsync(script));
        var result = document.RootElement;
        return new(result.GetProperty("detected").GetBoolean(), result.GetProperty("user").GetBoolean(),
            result.GetProperty("password").GetBoolean(), result.GetProperty("submitted").GetBoolean());
    }

    private sealed record LoginResult(bool Detected, bool User, bool Password, bool Submitted);
}
