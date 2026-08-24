using System.ComponentModel;
using System.Text.Json;
using System.Windows.Controls;
using Microsoft.Web.WebView2.Core;
using TaskTool.Services;
using TaskTool.ViewModels;

namespace TaskTool.Views;

public partial class WikiBrowserView : UserControl
{
    private WikiBrowserViewModel? _viewModel;
    public WikiBrowserView()
    {
        InitializeComponent(); DataContextChanged += OnDataContextChanged;
    }

    private void OnDataContextChanged(object sender, System.Windows.DependencyPropertyChangedEventArgs e)
    {
        if (_viewModel != null) { _viewModel.PropertyChanged -= OnPropertyChanged; _viewModel.BackRequested -= GoBack; _viewModel.ForwardRequested -= GoForward; _viewModel.ReloadRequested -= Reload; }
        _viewModel = e.NewValue as WikiBrowserViewModel;
        if (_viewModel != null) { _viewModel.PropertyChanged += OnPropertyChanged; _viewModel.BackRequested += GoBack; _viewModel.ForwardRequested += GoForward; _viewModel.ReloadRequested += Reload; _viewModel.EnsureHome(); _ = NavigateAsync(); }
    }
    private void OnPropertyChanged(object? sender, PropertyChangedEventArgs e) { if (e.PropertyName == nameof(WikiBrowserViewModel.NavigationUrl)) _ = NavigateAsync(); }
    private void GoBack() { if (WikiBrowser.CanGoBack) WikiBrowser.GoBack(); }
    private void GoForward() { if (WikiBrowser.CanGoForward) WikiBrowser.GoForward(); }
    private void Reload() => WikiBrowser.Reload();

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
        if (!Uri.TryCreate(e.Uri, UriKind.Absolute, out var uri) || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps)) { e.Cancel = true; BrowserStatus.Text = "Navigation wurde aus Sicherheitsgründen blockiert."; }
    }

    private async void OnNavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
    {
        var source = _viewModel?.SelectedSource;
        if (!e.IsSuccess || source?.BrowserLoginMode != "UsernamePassword" || !Uri.TryCreate(WikiBrowser.Source?.ToString(), UriKind.Absolute, out var current) || !Uri.TryCreate(WikiBrowserViewModel.GetHomeUrl(source), UriKind.Absolute, out var trusted) || current.Scheme != Uri.UriSchemeHttps || trusted.Scheme != Uri.UriSchemeHttps || !string.Equals(current.Scheme, trusted.Scheme, StringComparison.OrdinalIgnoreCase) || !string.Equals(current.Host, trusted.Host, StringComparison.OrdinalIgnoreCase) || current.Port != trusted.Port) return;
        var password = ServiceLocator.Settings.GetWikiBrowserPassword(source); if (string.IsNullOrWhiteSpace(source.BrowserUsername) || string.IsNullOrEmpty(password)) return;
        var userJson = JsonSerializer.Serialize(source.BrowserUsername); var passwordJson = JsonSerializer.Serialize(password); var submit = source.BrowserAutoSubmit ? "true" : "false";
        var script = $$"""(() => { const p=document.querySelector('input[type="password"]'); const u=document.querySelector('input[autocomplete="username"],input[name="username"],input[name="user"],input[type="email"]'); const f=p?.closest('form'); if(!p||!u||!f)return false; u.value={{userJson}};p.value={{passwordJson}};u.dispatchEvent(new Event('input',{bubbles:true}));p.dispatchEvent(new Event('input',{bubbles:true}));if({{submit}})f.querySelector('button[type="submit"],input[type="submit"]')?.click();return true;})()""";
        try { await WikiBrowser.CoreWebView2.ExecuteScriptAsync(script); } catch { BrowserStatus.Text = "Automatisches Ausfüllen war auf dieser Login-Seite nicht möglich."; }
    }
}
