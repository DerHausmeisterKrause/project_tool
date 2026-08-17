using System.Windows.Controls;
using System.Windows;
using System.Windows.Navigation;
using TaskTool.Services;

namespace TaskTool.Views;

public partial class TaskDetailPanel : UserControl
{
    public TaskDetailPanel()
    {
        InitializeComponent();
    }

    private void TaskDetailPanel_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        var stacked = e.NewSize.Width < 620;
        BookingTimerColumn.Width = new GridLength(stacked ? 1 : 1.2, GridUnitType.Star);
        BookingHistoryColumn.Width = stacked ? new GridLength(0) : new GridLength(1, GridUnitType.Star);
        Grid.SetColumn(BookingHistoryPanel, stacked ? 0 : 1);
        Grid.SetRow(BookingHistoryPanel, stacked ? 1 : 0);
        BookingTimerPanel.Margin = stacked ? new Thickness(0, 0, 0, 12) : new Thickness(0, 0, 12, 0);
    }

    private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
    {
        ServiceLocator.MainViewModel.NavigateToTicketSystem(e.Uri.AbsoluteUri);
        e.Handled = true;
    }
}
