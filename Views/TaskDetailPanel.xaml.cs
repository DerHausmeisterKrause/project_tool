using System.Diagnostics;
using System.Windows.Controls;
using System.Windows.Navigation;

namespace TaskTool.Views;

public partial class TaskDetailPanel : UserControl
{
    public TaskDetailPanel()
    {
        InitializeComponent();
    }

    private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
    {
        try
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri) { UseShellExecute = true });
        }
        catch
        {
            // no-op to keep UI stable when invalid URL is entered
        }
    }
}
