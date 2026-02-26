using System.Collections.ObjectModel;
using TaskTool.Infrastructure;
using TaskTool.Services;

namespace TaskTool.ViewModels;

public class MainViewModel : ObservableObject
{
    public ObservableCollection<object> NavigationItems { get; }
    public TodayViewModel TodayViewModel { get; }
    private readonly WeekViewModel _weekViewModel;

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
                Raise(nameof(IsTodaySelected));
            }
        }
    }

    public bool IsTodaySelected => SelectedView is TodayViewModel;

    public MainViewModel(TaskService taskService, WorkDayService workDayService, SettingsService settingsService, NotificationService notifications, LoggerService logger)
    {
        TodayViewModel = new TodayViewModel(taskService, workDayService, settingsService);
        _weekViewModel = new WeekViewModel(taskService, workDayService, settingsService);
        var reports = new ReportsViewModel(taskService, workDayService, settingsService);
        var settings = new SettingsViewModel(settingsService);

        NavigationItems = new ObservableCollection<object> { TodayViewModel, _weekViewModel, reports, settings };
        _selectedView = TodayViewModel;
    }
}
