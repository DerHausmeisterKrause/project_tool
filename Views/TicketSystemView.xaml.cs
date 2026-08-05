using System.ComponentModel;
using System.Windows.Controls;
using TaskTool.ViewModels;

namespace TaskTool.Views;

public partial class TicketSystemView : UserControl
{
    public TicketSystemView()
    {
        InitializeComponent();
        DataContextChanged += (_, _) => _ = NavigateToConfiguredUrlAsync();
    }

    protected override void OnInitialized(EventArgs e)
    {
        base.OnInitialized(e);
        if (DataContext is INotifyPropertyChanged notify)
            notify.PropertyChanged += OnViewModelPropertyChanged;
        _ = NavigateToConfiguredUrlAsync();
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(TicketSystemViewModel.TicketSystemWebUrl))
            _ = NavigateToConfiguredUrlAsync();
    }

    private async Task NavigateToConfiguredUrlAsync()
    {
        if (DataContext is not TicketSystemViewModel vm)
            return;

        if (!Uri.TryCreate(vm.TicketSystemWebUrl, UriKind.Absolute, out var uri))
        {
            BrowserStatus.Text = "Bitte in den Einstellungen eine gültige Ticketsystem-Webseite URL hinterlegen.";
            return;
        }

        try
        {
            BrowserStatus.Text = string.Empty;
            await TicketBrowser.EnsureCoreWebView2Async();
            TicketBrowser.CoreWebView2.Navigate(uri.ToString());
        }
        catch (Exception ex)
        {
            BrowserStatus.Text = $"WebView2 konnte nicht gestartet werden: {ex.Message}. Bitte Microsoft Edge WebView2 Runtime installieren/aktualisieren.";
        }
    }
}
