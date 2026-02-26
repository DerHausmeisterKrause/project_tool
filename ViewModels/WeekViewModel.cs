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
    public RelayCommand<WeekCalendarItem> OpenCalendarItemCommand { get; }
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
        OpenCalendarItemCommand = new RelayCommand<WeekCalendarItem>(OpenCalendarItem, i => i != null);
        OpenTicketUrlCommand = new RelayCommand<string>(OpenTicketUrlFromWeek, url => !string.IsNullOrWhiteSpace(url));
        SetDayTypeNormalCommand = new RelayCommand(() => SetDayType("Normal"), () => SelectedDay != null);
        SetDayTypeUlCommand = new RelayCommand(() => SetDayType("UL"), () => SelectedDay != null);
        SetDayTypeAmCommand = new RelayCommand(() => SetDayType("AM"), () => SelectedDay != null);
        ToggleHoCommand = new RelayCommand(() => { if (SelectedDay == null) return; SelectedDay.IsHo = !SelectedDay.IsHo; SaveSelectedDay(); }, () => SelectedDay != null);
        ToggleBrCommand = new RelayCommand(() => { if (SelectedDay == null) return; SelectedDay.IsBr = !SelectedDay.IsBr; SaveSelectedDay(); }, () => SelectedDay != null);

        LoadWeek();
    }

    private void OpenCalendarItem(WeekCalendarItem? item)
    {
        if (item == null) return;

        var main = ServiceLocator.MainViewModel;
        main.SelectedView = main.TodayViewModel;

        var match = main.TodayViewModel.CurrentTasks.FirstOrDefault(t => t.Id == item.TaskId)
                    ?? main.TodayViewModel.CompletedTasks.FirstOrDefault(t => t.Id == item.TaskId)
                    ?? _tasks.GetAllTasks().FirstOrDefault(t => t.Id == item.TaskId);

        if (match != null)
            main.TodayViewModel.SelectedTask = match;
    }

    private void OpenTicketUrlFromWeek(string? url)
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
        var from = WeekStart.Date;
        var to = WeekStart.AddDays(7).Date;

        var workDays = _workDays.GetWorkDaysInRange(from, to.AddDays(-1)).ToDictionary(w => w.Day, w => w);
        var segmentsInWeek = _tasks.GetSegmentsForRange(from, to)
            .GroupBy(x => x.Segment.StartLocal.Date)
            .ToDictionary(g => g.Key, g => g.OrderBy(x => x.Segment.StartLocal).ToList());

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

            var calendarItems = new List<WeekCalendarItem>();
            if (segmentsInWeek.TryGetValue(day.Date, out var segmentItems))
            {
                var indexByTask = new Dictionary<Guid, int>();
                foreach (var pair in segmentItems)
                {
                    indexByTask[pair.Task.Id] = indexByTask.TryGetValue(pair.Task.Id, out var current) ? current + 1 : 1;
                    calendarItems.Add(new WeekCalendarItem
                    {
                        TaskId = pair.Task.Id,
                        TaskTitle = pair.Task.Title,
                        TicketUrl = pair.Task.TicketUrl,
                        SegmentId = pair.Segment.Id,
                        SegmentStart = pair.Segment.StartLocal,
                        SegmentEnd = pair.Segment.EndLocal,
                        SegmentIndexDisplay = indexByTask[pair.Task.Id],
                        TaskStatus = pair.Task.Status.ToString()
                    });
                }
            }
            else
            {
                var fallbackTasks = _tasks.GetAllTasks()
                    .Where(t => t.StartLocal.HasValue && t.StartLocal.Value.Date == day.Date)
                    .OrderBy(t => t.StartLocal)
                    .ToList();

                var idx = 1;
                foreach (var task in fallbackTasks)
                {
                    calendarItems.Add(new WeekCalendarItem
                    {
                        TaskId = task.Id,
                        TaskTitle = task.Title,
                        TicketUrl = task.TicketUrl,
                        SegmentId = 0,
                        SegmentStart = task.StartLocal ?? day,
                        SegmentEnd = task.EndLocal ?? task.StartLocal ?? day,
                        SegmentIndexDisplay = idx++,
                        TaskStatus = task.Status.ToString()
                    });
                }
            }

            Days.Add(new WeekDayGroup
            {
                DayDate = day,
                DayLabel = day.ToString("ddd dd.MM"),
                CalendarItems = new ObservableCollection<WeekCalendarItem>(calendarItems),
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
            return Days.FirstOrDefault(d => d.DayDate.Date == today.Date) ?? Days.First();

        return Days.First();
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
    public ObservableCollection<WeekCalendarItem> CalendarItems { get; set; } = new();

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

public class WeekCalendarItem
{
    public Guid TaskId { get; set; }
    public string TaskTitle { get; set; } = string.Empty;
    public string TicketUrl { get; set; } = string.Empty;
    public long SegmentId { get; set; }
    public DateTime SegmentStart { get; set; }
    public DateTime SegmentEnd { get; set; }
    public int SegmentIndexDisplay { get; set; }
    public string TaskStatus { get; set; } = string.Empty;
}
