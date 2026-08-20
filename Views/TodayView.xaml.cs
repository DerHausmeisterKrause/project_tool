using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using TaskTool.Models;
using TaskTool.ViewModels;
using TaskTool.Services;

namespace TaskTool.Views;

public partial class TodayView : UserControl
{
    public TodayView()
    {
        InitializeComponent();
        DataContextChanged += OnDataContextChanged;
    }

    private void WorkDayHeaderGrid_SizeChanged(object sender, SizeChangedEventArgs e)
    {
        var compact = e.NewSize.Width < 620;
        Grid.SetRow(WorkDayMarkerPanel, compact ? 1 : 0);
        Grid.SetColumn(WorkDayMarkerPanel, compact ? 0 : 1);
        Grid.SetColumnSpan(WorkDayMarkerPanel, compact ? 2 : 1);
        WorkDayMarkerPanel.HorizontalAlignment = compact ? HorizontalAlignment.Left : HorizontalAlignment.Right;
        WorkDayMarkerPanel.Margin = compact ? new Thickness(0, 12, 0, 0) : new Thickness(0);
        WorkDayStatisticsPanel.Margin = compact ? new Thickness(0) : new Thickness(0, 0, 24, 0);
    }

    private void OnDataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
    {
        if (e.OldValue is TodayViewModel oldVm)
            oldVm.TaskBringIntoViewRequested -= OnTaskBringIntoViewRequested;

        var oldMain = ServiceLocator.MainViewModel;
        if (oldMain != null)
            oldMain.FocusQuickAddRequested -= OnFocusQuickAddRequested;

        if (e.NewValue is TodayViewModel vm)
            vm.TaskBringIntoViewRequested += OnTaskBringIntoViewRequested;

        ServiceLocator.MainViewModel.FocusQuickAddRequested += OnFocusQuickAddRequested;
    }

    private void OnTaskBringIntoViewRequested(Guid taskId)
    {
        Dispatcher.BeginInvoke(new Action(() =>
        {
            var element = FindTaskElement(TodayAgendaItems, taskId)
                          ?? FindTaskElement(CurrentTasksItems, taskId)
                          ?? FindTaskElement(CompletedTasksItems, taskId);
            element?.BringIntoView();
        }));
    }


    private void OnFocusQuickAddRequested()
    {
        Dispatcher.BeginInvoke(new Action(() =>
        {
            QuickAddTextBox.Focus();
            Keyboard.Focus(QuickAddTextBox);
            QuickAddTextBox.SelectAll();
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

            if (child is FrameworkElement agendaElement
                && agendaElement.DataContext is TodayAgendaItem agendaItem
                && agendaItem.TaskId == taskId)
                return agendaElement;

            var match = FindTaskElement(child, taskId);
            if (match != null)
                return match;
        }

        return null;
    }
}
