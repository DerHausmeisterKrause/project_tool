using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using TaskTool.ViewModels;

namespace TaskTool.Views;

public partial class WeekView : UserControl
{
    public WeekView()
    {
        InitializeComponent();
        Loaded += (_, _) => RefreshNowIndicator();
        SizeChanged += (_, _) => RefreshNowIndicator();
    }

    private void CalendarScrollViewer_ScrollChanged(object sender, ScrollChangedEventArgs e) => RefreshNowIndicator();

    private void RefreshNowIndicator()
    {
        if (DataContext is WeekViewModel vm)
            vm.UpdateTimelineMetrics(DaySegmentColumns.ActualHeight);
    }

    private void DayColumn_PreviewMouseDown(object sender, MouseButtonEventArgs e)
    {
        if (DataContext is not WeekViewModel vm) return;
        if (sender is FrameworkElement { DataContext: WeekDayGroup day })
            vm.SelectDayCommand.Execute(day);
    }
}
