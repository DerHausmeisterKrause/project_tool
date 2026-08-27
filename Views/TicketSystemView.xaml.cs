using System.ComponentModel;
using System.Text.Json;
using System.Windows.Controls;
using Microsoft.Web.WebView2.Core;
using TaskTool.Services;
using TaskTool.ViewModels;

namespace TaskTool.Views;

public partial class TicketSystemView : UserControl
{
    private INotifyPropertyChanged? _viewModelNotifications;
    private string _lastNavigationUrl = string.Empty;
    private Task? _initialization;

    public TicketSystemView()
    {
        InitializeComponent();
        DataContextChanged += (_, e) =>
        {
            if (_viewModelNotifications != null)
                _viewModelNotifications.PropertyChanged -= OnViewModelPropertyChanged;
            _viewModelNotifications = e.NewValue as INotifyPropertyChanged;
            if (_viewModelNotifications != null)
                _viewModelNotifications.PropertyChanged += OnViewModelPropertyChanged;
            _ = NavigateToConfiguredUrlAsync();
        };
    }

    protected override void OnInitialized(EventArgs e)
    {
        base.OnInitialized(e);
        if (_viewModelNotifications != null)
            _viewModelNotifications.PropertyChanged -= OnViewModelPropertyChanged;
        _viewModelNotifications = DataContext as INotifyPropertyChanged;
        if (_viewModelNotifications != null)
            _viewModelNotifications.PropertyChanged += OnViewModelPropertyChanged;
        // DataContextChanged owns first navigation; OnInitialized only wires the VM.
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(TicketSystemViewModel.NavigationUrl))
            _ = NavigateToConfiguredUrlAsync();
    }

    private async Task NavigateToConfiguredUrlAsync()
    {
        if (DataContext is not TicketSystemViewModel vm)
            return;

        if (!Uri.TryCreate(vm.NavigationUrl, UriKind.Absolute, out var uri))
        {
            BrowserStatus.Text = "Bitte in den Einstellungen eine gültige Ticketsystem-Webseite URL hinterlegen.";
            return;
        }

        try
        {
            var normalized = uri.AbsoluteUri;
            if (string.Equals(_lastNavigationUrl, normalized, StringComparison.OrdinalIgnoreCase)) return;
            BrowserStatus.Text = string.Empty;
            _initialization ??= TicketBrowser.EnsureCoreWebView2Async();
            await _initialization;
            TicketBrowser.CoreWebView2.NavigationCompleted -= TicketBrowser_NavigationCompleted;
            TicketBrowser.CoreWebView2.NavigationCompleted += TicketBrowser_NavigationCompleted;
            _lastNavigationUrl = normalized;
            TicketBrowser.CoreWebView2.Navigate(normalized);
        }
        catch (Exception ex)
        {
            BrowserStatus.Text = $"WebView2 konnte nicht gestartet werden: {ex.Message}. Bitte Microsoft Edge WebView2 Runtime installieren/aktualisieren.";
        }
    }

    private async void TicketBrowser_NavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
    {
        if (!e.IsSuccess || DataContext is not TicketSystemViewModel vm || !vm.AutofillCredentials)
            return;

        if (!Uri.TryCreate(TicketBrowser.Source?.ToString(), UriKind.Absolute, out var current)
            || !Uri.TryCreate(vm.TicketSystemWebUrl, UriKind.Absolute, out var configured)
            || current.Scheme != Uri.UriSchemeHttps
            || configured.Scheme != Uri.UriSchemeHttps
            || !string.Equals(current.Host, configured.Host, StringComparison.OrdinalIgnoreCase))
        {
            ServiceLocator.Logger.Info($"[TicketSystemLogin] host='{current?.Host ?? "invalid"}' loginPageDetected=False credentialsBlocked=True reason=UntrustedOrigin");
            return;
        }

        if (string.IsNullOrWhiteSpace(vm.Username) || string.IsNullOrEmpty(vm.Password))
            return;

        var usernameJson = JsonSerializer.Serialize(vm.Username);
        var passwordJson = JsonSerializer.Serialize(vm.Password);
        var autoSubmit = vm.AutoLogin ? "true" : "false";
        var script = $$"""
            (() => {
              const password = document.querySelector('input[name="Password"], input#Password, input[type="password"]');
              const username = document.querySelector('input[name="User"], input[name="UserLogin"], input#User, input#UserLogin, input[autocomplete="username"]');
              const loginForm = password && password.closest('form');
              const looksLikeLogin = !!password && !!username && !!loginForm;
              if (!looksLikeLogin) return { loginPageDetected:false, usernameFilled:false, passwordFilled:false, submitted:false };
              username.value = {{usernameJson}};
              password.value = {{passwordJson}};
              username.dispatchEvent(new Event('input', { bubbles:true }));
              password.dispatchEvent(new Event('input', { bubbles:true }));
              let submitted = false;
              if ({{autoSubmit}}) {
                const submit = loginForm.querySelector('button[type="submit"], input[type="submit"], button[name="LoginButton"], #LoginButton');
                if (submit) { submit.click(); submitted = true; }
              }
              return { loginPageDetected:true, usernameFilled:true, passwordFilled:true, submitted };
            })();
            """;

        try
        {
            var resultJson = await TicketBrowser.CoreWebView2.ExecuteScriptAsync(script);
            ServiceLocator.Logger.Info($"[TicketSystemLogin] host='{current.Host}' result={resultJson} autoSubmit={vm.AutoLogin}");
        }
        catch (Exception ex)
        {
            ServiceLocator.Logger.Error($"[TicketSystemLogin] host='{current.Host}' autofillFailed=True message='{ex.Message}'");
        }
    }
}
