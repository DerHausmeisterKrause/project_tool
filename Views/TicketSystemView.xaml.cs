using System.ComponentModel;
using System.Windows.Controls;
using TaskTool.ViewModels;

namespace TaskTool.Views;

public partial class TicketSystemView : UserControl
{
    public TicketSystemView()
    {
        InitializeComponent();
        DataContextChanged += (_, _) => NavigateToConfiguredUrl();
    }

    protected override void OnInitialized(EventArgs e)
    {
        base.OnInitialized(e);
        if (DataContext is INotifyPropertyChanged notify)
            notify.PropertyChanged += OnViewModelPropertyChanged;
        NavigateToConfiguredUrl();
    }

    private void OnViewModelPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(TicketSystemViewModel.TicketSystemWebUrl))
            NavigateToConfiguredUrl();
    }

    private void NavigateToConfiguredUrl()
    {
        if (DataContext is not TicketSystemViewModel vm)
            return;

        if (!Uri.TryCreate(vm.TicketSystemWebUrl, UriKind.Absolute, out var uri))
            return;

        TicketBrowser.Navigate(uri);
    }
}
