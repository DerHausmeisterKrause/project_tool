using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using TaskTool.Models;
using TaskTool.ViewModels;

namespace TaskTool.Views;

public partial class TodayView : UserControl
{
    public TodayView()
    {
        InitializeComponent();
        DataContextChanged += OnDataContextChanged;
    }

    private void OnDataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
    {
        if (e.OldValue is TodayViewModel oldVm)
            oldVm.TaskBringIntoViewRequested -= OnTaskBringIntoViewRequested;

        if (e.NewValue is TodayViewModel vm)
            vm.TaskBringIntoViewRequested += OnTaskBringIntoViewRequested;
    }

    private void OnTaskBringIntoViewRequested(Guid taskId)
    {
        Dispatcher.BeginInvoke(new Action(() =>
        {
            var element = FindTaskElement(CurrentTasksItems, taskId) ?? FindTaskElement(CompletedTasksItems, taskId);
            element?.BringIntoView();
        }));
    }

    private static FrameworkElement? FindTaskElement(DependencyObject root, Guid taskId)
    {
        var count = VisualTreeHelper.GetChildrenCount(root);
        for (var i = 0; i < count; i++)
        {
            var child = VisualTreeHelper.GetChild(root, i);
            if (child is FrameworkElement fe && fe.DataContext is TaskItem task && task.Id == taskId)
                return fe;

            var match = FindTaskElement(child, taskId);
            if (match != null)
                return match;
        }

        return null;
    }
}
