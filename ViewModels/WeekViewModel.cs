using System.Collections.ObjectModel;
using TaskTool.Infrastructure;
using TaskTool.Models;
using TaskTool.Services;

namespace TaskTool.ViewModels;

public class WeekViewModel : ObservableObject
{
    private readonly TaskService _tasks;
    private readonly WorkDayService _workDays;
    private readonly SettingsService _settings;

    public string Title => "Kalender";

    private DateTime _weekStart;
    public DateTime WeekStart
    {
        get => _weekStart;
        set { if (Set(ref _weekStart, value)) Raise(nameof(WeekRangeLabel)); }
    }

    public string WeekRangeLabel => $"{WeekStart:dd.MM.yyyy} - {WeekStart.AddDays(6):dd.MM.yyyy}";
    public ObservableCollection<WeekDayGroup> Days { get; } = new();

    private WeekDayGroup? _selectedDay;
    public WeekDayGroup? SelectedDay
    {
        get => _selectedDay;
        set
        {
            if (Set(ref _selectedDay, value))
            {
                foreach (var d in Days) d.IsSelected = d == value;
                Raise(nameof(SelectedDayType));
                Raise(nameof(SelectedIsHo));
                Raise(nameof(SelectedIsBr));
                SetDayTypeNormalCommand.RaiseCanExecuteChanged();
                SetDayTypeUlCommand.RaiseCanExecuteChanged();
                SetDayTypeAmCommand.RaiseCanExecuteChanged();
                ToggleHoCommand.RaiseCanExecuteChanged();
                ToggleBrCommand.RaiseCanExecuteChanged();
            }
        }
    }

    public string SelectedDayType => SelectedDay?.DayType ?? "Normal";
    public bool SelectedIsHo => SelectedDay?.IsHo ?? false;
    public bool SelectedIsBr => SelectedDay?.IsBr ?? false;

    public RelayCommand PreviousWeekCommand { get; }
    public RelayCommand NextWeekCommand { get; }
    public RelayCommand CurrentWeekCommand { get; }
    public RelayCommand<WeekDayGroup> SelectDayCommand { get; }
    public RelayCommand<TaskItem> OpenTaskCommand { get; }
    public RelayCommand<string> OpenTicketUrlCommand { get; }
    public RelayCommand SetDayTypeNormalCommand { get; }
    public RelayCommand SetDayTypeUlCommand { get; }
    public RelayCommand SetDayTypeAmCommand { get; }
    public RelayCommand ToggleHoCommand { get; }
    public RelayCommand ToggleBrCommand { get; }

    public WeekViewModel(TaskService tasks, WorkDayService workDays, SettingsService settings)
    {
        _tasks = tasks;
        _workDays = workDays;
        _settings = settings;

        WeekStart = StartOfWeek(DateTime.Today);
        PreviousWeekCommand = new RelayCommand(() => { WeekStart = WeekStart.AddDays(-7); LoadWeek(); });
        NextWeekCommand = new RelayCommand(() => { WeekStart = WeekStart.AddDays(7); LoadWeek(); });
        CurrentWeekCommand = new RelayCommand(() => { WeekStart = StartOfWeek(DateTime.Today); LoadWeek(); });
        SelectDayCommand = new RelayCommand<WeekDayGroup>(d => SelectedDay = d, d => d != null);
        OpenTaskCommand = new RelayCommand<TaskItem>(OpenTask, t => t != null);
        OpenTicketUrlCommand = new RelayCommand<string>(OpenTicketUrl, url => !string.IsNullOrWhiteSpace(url));
        SetDayTypeNormalCommand = new RelayCommand(() => SetDayType("Normal"), () => SelectedDay != null);
        SetDayTypeUlCommand = new RelayCommand(() => SetDayType("UL"), () => SelectedDay != null);
        SetDayTypeAmCommand = new RelayCommand(() => SetDayType("AM"), () => SelectedDay != null);
        ToggleHoCommand = new RelayCommand(() => { if (SelectedDay == null) return; SelectedDay.IsHo = !SelectedDay.IsHo; SaveSelectedDay(); }, () => SelectedDay != null);
        ToggleBrCommand = new RelayCommand(() => { if (SelectedDay == null) return; SelectedDay.IsBr = !SelectedDay.IsBr; SaveSelectedDay(); }, () => SelectedDay != null);

        LoadWeek();
    }

    private void OpenTask(TaskItem? task)
    {
        if (task == null) return;
        var main = ServiceLocator.MainViewModel;
        main.SelectedView = main.TodayViewModel;
        var match = main.TodayViewModel.CurrentTasks.FirstOrDefault(t => t.Id == task.Id)
                    ?? main.TodayViewModel.CompletedTasks.FirstOrDefault(t => t.Id == task.Id)
                    ?? task;
        main.TodayViewModel.SelectedTask = match;
    }

    private void OpenTicketUrl(string? url)
    {
        UrlLauncher.TryOpen(url, out _);
    }

    private void SetDayType(string type)
    {
        if (SelectedDay == null) return;
        SelectedDay.DayType = type;
        SaveSelectedDay();
    }

    private void SaveSelectedDay()
    {
        if (SelectedDay == null) return;
        _workDays.SetDayMarkers(SelectedDay.DayDate.ToString("yyyy-MM-dd"), SelectedDay.DayType, SelectedDay.IsBr, SelectedDay.IsHo);
        var selectedDate = SelectedDay.DayDate;
        LoadWeek();
        SelectedDay = Days.FirstOrDefault(d => d.DayDate.Date == selectedDate.Date) ?? Days.FirstOrDefault();
    }

    private void LoadWeek()
    {
        var previousSelectionDate = SelectedDay?.DayDate;

        Days.Clear();
        var from = WeekStart;
        var to = WeekStart.AddDays(7);

        var allTasks = _tasks.GetAllTasks();
        var workDays = _workDays.GetWorkDaysInRange(from, to.AddDays(-1)).ToDictionary(w => w.Day, w => w);

        for (int i = 0; i < 7; i++)
        {
            var day = WeekStart.AddDays(i);
            var key = day.ToString("yyyy-MM-dd");
            if (!workDays.ContainsKey(key)) workDays[key] = _workDays.GetOrCreateDay(key);

            var wd = workDays[key];
            var breaks = _workDays.GetBreaks(key);
            var pause = breaks.Where(b => b.EndLocal.HasValue).Sum(b => (int)(b.EndLocal!.Value - b.StartLocal).TotalMinutes);
            var net = (wd.ComeLocal.HasValue && wd.GoLocal.HasValue) ? (int)(wd.GoLocal.Value - wd.ComeLocal.Value).TotalMinutes - pause : 0;
            var target = (wd.DayType == "UL" || wd.DayType == "AM") ? 0 : _settings.Current.GetTargetMinutes(day.DayOfWeek);
            var overtime = net - target;

            var dayTasks = allTasks
                .Where(t => IsTaskOnDay(t, day))
                .OrderBy(t => t.StartLocal ?? t.CreatedUtc)
                .ToList();

            Days.Add(new WeekDayGroup
            {
                DayDate = day,
                DayLabel = day.ToString("ddd dd.MM"),
                Tasks = new ObservableCollection<TaskItem>(dayTasks),
                DayType = wd.DayType,
                IsBr = wd.IsBr,
                IsHo = wd.IsHo,
                Summary = $"Soll {Fmt(target)} | Ist {Fmt(net)} | Ü {Fmt(overtime)}"
            });
        }

        SelectedDay = ResolveSelectedDay(previousSelectionDate);
    }

    private WeekDayGroup ResolveSelectedDay(DateTime? previousSelectionDate)
    {
        if (previousSelectionDate.HasValue)
        {
            var existing = Days.FirstOrDefault(d => d.DayDate.Date == previousSelectionDate.Value.Date);
            if (existing != null) return existing;
        }

        var today = DateTime.Today;
        var containsToday = today.Date >= WeekStart.Date && today.Date <= WeekStart.AddDays(6).Date;
        if (containsToday)
        {
            return Days.FirstOrDefault(d => d.DayDate.Date == today.Date) ?? Days.First();
        }

        return Days.First();
    }

    private static bool IsTaskOnDay(TaskItem task, DateTime day)
    {
        if (!task.StartLocal.HasValue) return false;

        var dayStart = day.Date;
        var dayEnd = dayStart.AddDays(1);

        var start = task.StartLocal.Value;
        var end = task.EndLocal ?? start;

        return start < dayEnd && end >= dayStart;
    }

    private static string Fmt(int minutes) => $"{minutes / 60}h {Math.Abs(minutes % 60):00}m";
    private static DateTime StartOfWeek(DateTime date) => date.Date.AddDays(-((7 + (date.DayOfWeek - DayOfWeek.Monday)) % 7));

    public void Refresh() => LoadWeek();

    public override string ToString() => Title;
}

public class WeekDayGroup : ObservableObject
{
    public DateTime DayDate { get; set; }
    public string DayLabel { get; set; } = string.Empty;
    public ObservableCollection<TaskItem> Tasks { get; set; } = new();

    private string _dayType = "Normal";
    public string DayType { get => _dayType; set => Set(ref _dayType, value); }

    private bool _isBr;
    public bool IsBr { get => _isBr; set => Set(ref _isBr, value); }

    private bool _isHo;
    public bool IsHo { get => _isHo; set => Set(ref _isHo, value); }

    private string _summary = string.Empty;
    public string Summary { get => _summary; set => Set(ref _summary, value); }

    private bool _isSelected;
    public bool IsSelected { get => _isSelected; set => Set(ref _isSelected, value); }
}
