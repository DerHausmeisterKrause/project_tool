using System.Windows;
using System.Windows.Controls;
using TaskTool.Services;
using TaskTool.ViewModels;

namespace TaskTool.Views;

public partial class WebShortcutView : UserControl
{
    private WebShortcutBrowserSession? _session;
    private long _attachGeneration;

    public WebShortcutView()
    {
        InitializeComponent();
        Loaded += OnLoaded;
        Unloaded += OnUnloaded;
        DataContextChanged += OnDataContextChanged;
    }

    private void OnLoaded(object sender, RoutedEventArgs e) => _ = AttachCurrentAsync();

    private void OnUnloaded(object sender, RoutedEventArgs e)
    {
        _attachGeneration++;
        DetachSession();
    }

    private void OnDataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
    {
        _attachGeneration++;
        DetachSession();
        if (IsLoaded) _ = AttachCurrentAsync();
    }

    private async Task AttachCurrentAsync()
    {
        if (DataContext is not WebShortcutViewModel viewModel) return;
        var shortcutId = viewModel.ShortcutId;
        var generation = ++_attachGeneration;
        BrowserStatus.Text = "Webseite wird geladen …";
        try
        {
            var session = ServiceLocator.WebShortcutBrowsers.GetOrCreate(viewModel.Shortcut);
            _session = session;
            _session.StatusChanged += OnStatusChanged;
            _session.AttachTo(BrowserHost);
            await session.EnsureInitializedAsync();
            if (generation != _attachGeneration
                || !IsLoaded
                || DataContext is not WebShortcutViewModel current
                || !string.Equals(current.ShortcutId, shortcutId, StringComparison.Ordinal)) return;
            BrowserStatus.Text = _session.LastStatus;
        }
        catch (Exception ex)
        {
            if (generation == _attachGeneration) BrowserStatus.Text = $"Webseite konnte nicht geöffnet werden: {ex.Message}";
        }
    }

    private void OnStatusChanged(string status)
    {
        if (Dispatcher.CheckAccess()) BrowserStatus.Text = status;
        else Dispatcher.Invoke(() => BrowserStatus.Text = status);
    }

    private void DetachSession()
    {
        if (_session == null) return;
        _session.StatusChanged -= OnStatusChanged;
        _session.DetachFrom(BrowserHost);
        _session = null;
    }
}
