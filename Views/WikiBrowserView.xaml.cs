using System.ComponentModel;
using System.IO;
using System.Text.Json;
using System.Windows.Controls;
using Microsoft.Web.WebView2.Core;
using TaskTool.Services;
using TaskTool.ViewModels;

namespace TaskTool.Views;

public partial class WikiBrowserView : UserControl
{
    private WikiBrowserViewModel? _viewModel;
    private CancellationTokenSource? _loginAttemptCancellation;
    public WikiBrowserView()
    {
        InitializeComponent(); DataContextChanged += OnDataContextChanged;
    }

    private void OnDataContextChanged(object sender, System.Windows.DependencyPropertyChangedEventArgs e)
    {
        _loginAttemptCancellation?.Cancel();
        if (_viewModel != null) _viewModel.PropertyChanged -= OnPropertyChanged;
        _viewModel = e.NewValue as WikiBrowserViewModel;
        if (_viewModel != null) { _viewModel.PropertyChanged += OnPropertyChanged; _viewModel.EnsureHome(); _ = NavigateAsync(); }
    }
    private void OnPropertyChanged(object? sender, PropertyChangedEventArgs e) { if (e.PropertyName == nameof(WikiBrowserViewModel.NavigationUrl)) _ = NavigateAsync(); }

    private async Task NavigateAsync()
    {
        if (_viewModel == null || !Uri.TryCreate(_viewModel.NavigationUrl, UriKind.Absolute, out var uri) || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps)) return;
        try
        {
            var dataPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Plenaro", "WebView2", "Wiki");
            var environment = await CoreWebView2Environment.CreateAsync(null, dataPath);
            await WikiBrowser.EnsureCoreWebView2Async(environment);
            WikiBrowser.CoreWebView2.NavigationStarting -= OnNavigationStarting; WikiBrowser.CoreWebView2.NavigationStarting += OnNavigationStarting;
            WikiBrowser.CoreWebView2.NavigationCompleted -= OnNavigationCompleted; WikiBrowser.CoreWebView2.NavigationCompleted += OnNavigationCompleted;
            WikiBrowser.CoreWebView2.Navigate(uri.ToString()); BrowserStatus.Text = string.Empty;
        }
        catch (Exception ex) { BrowserStatus.Text = $"Wiki-Browser konnte nicht gestartet werden: {ex.Message}"; }
    }

    private void OnNavigationStarting(object? sender, CoreWebView2NavigationStartingEventArgs e)
    {
        _loginAttemptCancellation?.Cancel();
        if (!Uri.TryCreate(e.Uri, UriKind.Absolute, out var uri) || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps)) { e.Cancel = true; BrowserStatus.Text = "Navigation wurde aus Sicherheitsgründen blockiert."; }
    }

    private void OnNavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
    {
        var source = _viewModel?.SelectedSource;
        if (!e.IsSuccess || source?.BrowserLoginMode != "UsernamePassword" || !IsTrustedOrigin(source, out var host)) return;
        var password = ServiceLocator.Settings.GetWikiBrowserPassword(source);
        if (string.IsNullOrWhiteSpace(source.BrowserUsername) || string.IsNullOrEmpty(password)) return;
        _loginAttemptCancellation?.Cancel(); _loginAttemptCancellation = new CancellationTokenSource();
        _ = TryLoginWithRetriesAsync(source, host, password, _loginAttemptCancellation.Token);
    }

    private bool IsTrustedOrigin(TaskTool.Models.WikiSourceSettings source, out string host)
    {
        host = "invalid";
        if (!Uri.TryCreate(WikiBrowser.Source?.ToString(), UriKind.Absolute, out var current) || !Uri.TryCreate(source.BaseUrl, UriKind.Absolute, out var trusted)) return false;
        host = current.Host;
        return current.Scheme == Uri.UriSchemeHttps && trusted.Scheme == Uri.UriSchemeHttps && string.Equals(current.Host, trusted.Host, StringComparison.OrdinalIgnoreCase) && current.Port == trusted.Port;
    }

    private async Task TryLoginWithRetriesAsync(TaskTool.Models.WikiSourceSettings source, string host, string password, CancellationToken token)
    {
        var delays = new[] { 0, 300, 800, 1500 }; var previousDelay = 0;
        foreach (var delay in delays)
        {
            try
            {
                if (delay > previousDelay) await Task.Delay(delay - previousDelay, token);
                previousDelay = delay;
                token.ThrowIfCancellationRequested();
                var result = await TryFillAndSubmitLoginAsync(source.BrowserUsername, password);
                ServiceLocator.Logger.Info($"[WikiBrowserLogin] sourceId={source.Id} host={host} loginPageDetected={result.LoginPageDetected.ToString().ToLowerInvariant()} usernameFieldFound={result.UsernameFieldFound.ToString().ToLowerInvariant()} passwordFieldFound={result.PasswordFieldFound.ToString().ToLowerInvariant()} submitFound={result.SubmitFound.ToString().ToLowerInvariant()} submitted={result.Submitted.ToString().ToLowerInvariant()}");
                if (result.LoginPageDetected) return;
            }
            catch (OperationCanceledException) { return; }
            catch (Exception ex) { ServiceLocator.Logger.Warning($"[WikiBrowserLogin] sourceId={source.Id} host={host} status=failed errorType={ex.GetType().Name}"); return; }
        }
    }

    private async Task<LoginAttemptResult> TryFillAndSubmitLoginAsync(string username, string password)
    {
        var usernameJson = JsonSerializer.Serialize(username); var passwordJson = JsonSerializer.Serialize(password);
        var script = $$"""
            (() => {
              const first = selectors => selectors.map(s => document.querySelector(s)).find(Boolean) || null;
              const username = first(['input#os_username','input[name="os_username"]','input#username','input[name="username"]','input[name="user"]','input[name="User"]','input[autocomplete="username"]','input[type="email"]']);
              const password = first(['input#os_password','input[name="os_password"]','input#password','input[name="password"]','input[name="Password"]','input[type="password"]']);
              const form = password?.closest('form') || null;
              const submit = form && ['button#loginButton','input#loginButton','button[name="login"]','input[name="login"]','button[type="submit"]','input[type="submit"]'].map(s => form.querySelector(s)).find(Boolean) || null;
              const result = { loginPageDetected:!!(username && password), usernameFieldFound:!!username, passwordFieldFound:!!password, submitFound:!!submit, submitted:false };
              if (!username || !password) return result;
              const setValue = (element, value) => { element.focus(); const setter=Object.getOwnPropertyDescriptor(HTMLInputElement.prototype,'value')?.set; if(setter) setter.call(element,value); else element.value=value; element.dispatchEvent(new Event('input',{bubbles:true})); element.dispatchEvent(new Event('change',{bubbles:true})); };
              setValue(username, {{usernameJson}}); setValue(password, {{passwordJson}});
              if (submit) { submit.click(); result.submitted=true; }
              else if (form?.requestSubmit) { form.requestSubmit(); result.submitted=true; }
              return result;
            })()
            """;
        var json = await WikiBrowser.CoreWebView2.ExecuteScriptAsync(script);
        using var document = JsonDocument.Parse(json); var root = document.RootElement;
        return new LoginAttemptResult(GetBoolean(root, "loginPageDetected"), GetBoolean(root, "usernameFieldFound"), GetBoolean(root, "passwordFieldFound"), GetBoolean(root, "submitFound"), GetBoolean(root, "submitted"));
    }

    private static bool GetBoolean(JsonElement root, string name) => root.TryGetProperty(name, out var value) && value.ValueKind == JsonValueKind.True;
    private sealed record LoginAttemptResult(bool LoginPageDetected, bool UsernameFieldFound, bool PasswordFieldFound, bool SubmitFound, bool Submitted);
}
