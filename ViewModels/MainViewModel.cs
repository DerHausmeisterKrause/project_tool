using System;
using System.Windows;
using System.Collections.ObjectModel;
using TaskTool.Infrastructure;
using TaskTool.Services;

namespace TaskTool.ViewModels;

public class MainViewModel : ObservableObject
{
    public ObservableCollection<object> NavigationItems { get; }
    public TodayViewModel TodayViewModel { get; }
    private readonly WeekViewModel _weekViewModel;
    private readonly TicketSystemViewModel _ticketSystemViewModel;

    private object _selectedView;
    public object SelectedView
    {
        get => _selectedView;
        set
        {
            if (Set(ref _selectedView, value))
            {
                if (_selectedView is WeekViewModel)
                {
                    _weekViewModel.Refresh();
                }
                else if (_selectedView is TicketSystemViewModel)
                {
                    _ticketSystemViewModel.Refresh();
                }
                Raise(nameof(IsTodaySelected));
            }
        }
    }

    public bool IsTodaySelected => SelectedView is TodayViewModel;

    public event Action? FocusQuickAddRequested;

    public void NavigateToTodayAndOpenTask(Guid taskId)
    {
        SelectedView = TodayViewModel;

        if (TodayViewModel.NavigateToTask(taskId))
            return;

        MessageBox.Show("Aufgabe nicht gefunden", "Kalender", MessageBoxButton.OK, MessageBoxImage.Warning);
    }

    public void NavigateToTodayAndFocusQuickAdd()
    {
        SelectedView = TodayViewModel;
        FocusQuickAddRequested?.Invoke();
    }

    public MainViewModel(TaskService taskService, WorkDayService workDayService, SettingsService settingsService, NotificationService notifications, OutlookCalendarService outlookCalendar, TicketSystemService ticketSystem, LoggerService logger)
    {
        TodayViewModel = new TodayViewModel(taskService, workDayService, settingsService, outlookCalendar);
        _weekViewModel = new WeekViewModel(taskService, workDayService, settingsService, outlookCalendar);
        _ticketSystemViewModel = new TicketSystemViewModel(settingsService);
        var reports = new ReportsViewModel(taskService, workDayService, settingsService);
        var settings = new SettingsViewModel(settingsService, notifications, outlookCalendar, taskService, ticketSystem, TodayViewModel.Refresh);

        NavigationItems = new ObservableCollection<object> { TodayViewModel, _weekViewModel, _ticketSystemViewModel, reports, settings };
        _selectedView = TodayViewModel;
    }
}
