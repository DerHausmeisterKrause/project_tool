using System.Windows.Controls;
using System.Windows.Navigation;
using TaskTool.Services;

namespace TaskTool.Views;

public partial class TaskDetailPanel : UserControl
{
    public TaskDetailPanel()
    {
        InitializeComponent();
    }

    private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
    {
        ServiceLocator.MainViewModel.NavigateToTicketSystem(e.Uri.AbsoluteUri);
        e.Handled = true;
    }
}
