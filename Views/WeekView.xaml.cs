using System.Windows.Controls;
using System.Windows.Input;
using TaskTool.ViewModels;

namespace TaskTool.Views;

public partial class WeekView : UserControl
{
    public WeekView()
    {
        InitializeComponent();
    }

    private void DayColumn_PreviewMouseDown(object sender, MouseButtonEventArgs e)
    {
        if (DataContext is not WeekViewModel vm) return;
        if (sender is FrameworkElement { DataContext: WeekDayGroup day })
            vm.SelectDayCommand.Execute(day);
    }
}
