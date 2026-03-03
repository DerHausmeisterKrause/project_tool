using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows.Threading;
using TaskTool.Infrastructure;
using TaskTool.Models;
using TaskTool.Services;

namespace TaskTool.ViewModels;

public class WeekViewModel : ObservableObject
{
    private readonly TaskService _tasks;
    private readonly WorkDayService _workDays;
    private readonly SettingsService _settings;

    private const int CalendarStartHour = 6;
    private const int CalendarEndHour = 18;
    private const double PixelsPerMinuteConst = 1.2;
    private const double DayColumnWidthConst = 220;
    private const double DayInnerPadding = 4;
    private const double OverlapGap = 4;

    public string Title => "Kalender";

    public double TimeColumnWidth => 72;
    public double DayColumnWidth => DayColumnWidthConst;
    public double PixelsPerMinute => PixelsPerMinuteConst;
    public double CalendarBodyHeight => (CalendarEndHour - CalendarStartHour) * 60 * PixelsPerMinute;
    public double FullDayColumnHeight => CalendarBodyHeight + 58;
    public double DayAreaWidth => DayColumnWidth * 7;

    public ObservableCollection<TimeAxisLabel> TimeAxisLabels { get; } = new();
    public ObservableCollection<TimeGridLine> TimeGridLines { get; } = new();

    private readonly DispatcherTimer _nowIndicatorTimer = new();

    private bool _showNowIndicator;
    public bool ShowNowIndicator { get => _showNowIndicator; set => Set(ref _showNowIndicator, value); }

    private double _nowLineTop;
    public double NowLineTop { get => _nowLineTop; set => Set(ref _nowLineTop, value); }

    private double _nowMarkerLeft;
    public double NowMarkerLeft { get => _nowMarkerLeft; set => Set(ref _nowMarkerLeft, value); }

    private double _nowMarkerTop;
    public double NowMarkerTop { get => _nowMarkerTop; set => Set(ref _nowMarkerTop, value); }

    private DateTime _weekStart;
    public DateTime WeekStart
    {
        get => _weekStart;
        set { if (Set(ref _weekStart, value)) Raise(nameof(WeekRangeLabel)); }
    }

    public string WeekRangeLabel => $"{WeekStart:dd.MM.yyyy} - {WeekStart.AddDays(6):dd.MM.yyyy}";
    public ObservableCollection<WeekDayGroup> Days { get; } = new();

    private DateTime _selectedDate;
    public DateTime SelectedDate
    {
        get => _selectedDate;
        set
        {
            if (Set(ref _selectedDate, value.Date))
            {
                var day = Days.FirstOrDefault(d => d.DayDate.Date == _selectedDate.Date);
                if (day != null && !ReferenceEquals(day, SelectedDay))
                    SelectedDay = day;
            }
        }
    }

    private WeekDayGroup? _selectedDay;
    public WeekDayGroup? SelectedDay
    {
        get => _selectedDay;
        set
        {
            if (Set(ref _selectedDay, value))
            {
                if (value != null && value.DayDate.Date != SelectedDate.Date)
                    Set(ref _selectedDate, value.DayDate.Date, nameof(SelectedDate));

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

        BuildTimeScale();

        WeekStart = StartOfWeek(DateTime.Today);
        PreviousWeekCommand = new RelayCommand(() => { WeekStart = WeekStart.AddDays(-7); LoadWeek(); });
        NextWeekCommand = new RelayCommand(() => { WeekStart = WeekStart.AddDays(7); LoadWeek(); });
        CurrentWeekCommand = new RelayCommand(() => { WeekStart = StartOfWeek(DateTime.Today); LoadWeek(); });
        SelectDayCommand = new RelayCommand<WeekDayGroup>(d =>
        {
            if (d == null) return;
            SelectedDate = d.DayDate.Date;
        }, d => d != null);
        OpenCalendarItemCommand = new RelayCommand<WeekCalendarItem>(OpenCalendarItem, i => i != null);
        OpenTicketUrlCommand = new RelayCommand<string>(OpenTicketUrlFromWeek, url => !string.IsNullOrWhiteSpace(url));
        SetDayTypeNormalCommand = new RelayCommand(() => SetDayType("Normal"), () => SelectedDay != null);
        SetDayTypeUlCommand = new RelayCommand(() => SetDayType("UL"), () => SelectedDay != null);
        SetDayTypeAmCommand = new RelayCommand(() => SetDayType("AM"), () => SelectedDay != null);
        ToggleHoCommand = new RelayCommand(() => { if (SelectedDay == null) return; SelectedDay.IsHo = !SelectedDay.IsHo; SaveSelectedDay(); }, () => SelectedDay != null);
        ToggleBrCommand = new RelayCommand(() => { if (SelectedDay == null) return; SelectedDay.IsBr = !SelectedDay.IsBr; SaveSelectedDay(); }, () => SelectedDay != null);

        _nowIndicatorTimer.Interval = TimeSpan.FromSeconds(30);
        _nowIndicatorTimer.Tick += (_, _) => UpdateNowIndicator();
        _nowIndicatorTimer.Start();

        LoadWeek();
        UpdateNowIndicator();
    }

    private void BuildTimeScale()
    {
        TimeAxisLabels.Clear();
        TimeGridLines.Clear();

        var totalMinutes = (CalendarEndHour - CalendarStartHour) * 60;
        for (var minute = 0; minute <= totalMinutes; minute += 30)
        {
            var time = TimeSpan.FromHours(CalendarStartHour).Add(TimeSpan.FromMinutes(minute));
            var isHour = minute % 60 == 0;
            var top = minute * PixelsPerMinute;

            TimeGridLines.Add(new TimeGridLine { Top = top, IsHour = isHour });
            if (isHour)
            {
                TimeAxisLabels.Add(new TimeAxisLabel
                {
                    Label = $"{time.Hours:00}:00",
                    Top = top - 8
                });
            }
        }
    }

    private void OpenCalendarItem(WeekCalendarItem? item)
    {
        if (item == null) return;

        SelectedDate = item.SegmentStart.Date;

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
                        TaskDescription = pair.Task.Description,
                        TicketUrl = pair.Task.TicketUrl,
                        SegmentId = pair.Segment.Id,
                        SegmentStart = pair.Segment.StartLocal,
                        SegmentEnd = pair.Segment.EndLocal,
                        SegmentIndexDisplay = indexByTask[pair.Task.Id],
                        TaskStatus = pair.Task.Status.ToString(),
                        SegmentNote = pair.Segment.Note
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
                        TaskDescription = task.Description,
                        TicketUrl = task.TicketUrl,
                        SegmentId = 0,
                        SegmentStart = task.StartLocal ?? day,
                        SegmentEnd = task.EndLocal ?? task.StartLocal ?? day,
                        SegmentIndexDisplay = idx++,
                        TaskStatus = task.Status.ToString(),
                        SegmentNote = string.Empty
                    });
                }
            }

            LayoutDayItems(day.Date, calendarItems);

            Days.Add(new WeekDayGroup
            {
                DayDate = day,
                DayLabel = day.ToString("ddd dd.MM", CultureInfo.CurrentCulture),
                IsToday = day.Date == DateTime.Today,
                CalendarItems = new ObservableCollection<WeekCalendarItem>(calendarItems.OrderBy(c => c.DisplayTop)),
                DayType = wd.DayType,
                IsBr = wd.IsBr,
                IsHo = wd.IsHo,
                Summary = $"Soll {Fmt(target)} | Ist {Fmt(net)} | Ü {Fmt(overtime)}"
            });
        }

        SelectedDay = ResolveSelectedDay(previousSelectionDate);
        UpdateNowIndicator();
    }

    private void UpdateNowIndicator()
    {
        var now = DateTime.Now;
        var today = now.Date;
        var weekEnd = WeekStart.Date.AddDays(6);
        if (today < WeekStart.Date || today > weekEnd)
        {
            ShowNowIndicator = false;
            return;
        }

        var start = today.AddHours(CalendarStartHour);
        var minutesFromStart = (now - start).TotalMinutes;
        var rangeMinutes = (CalendarEndHour - CalendarStartHour) * 60;

        if (minutesFromStart < 0 || minutesFromStart > rangeMinutes)
        {
            ShowNowIndicator = false;
            return;
        }

        var dayIndex = (int)(today - WeekStart.Date).TotalDays;
        NowLineTop = minutesFromStart * PixelsPerMinute;
        NowMarkerLeft = dayIndex * DayColumnWidth - 4;
        NowMarkerTop = NowLineTop - 3;
        ShowNowIndicator = true;
    }

    private void LayoutDayItems(DateTime dayDate, List<WeekCalendarItem> items)
    {
        var rangeStart = dayDate.Date.AddHours(CalendarStartHour);
        var rangeEnd = dayDate.Date.AddHours(CalendarEndHour);

        foreach (var item in items)
        {
            item.DisplayStart = item.SegmentStart < rangeStart ? rangeStart : item.SegmentStart;
            item.DisplayEnd = item.SegmentEnd > rangeEnd ? rangeEnd : item.SegmentEnd;
        }

        var visible = items
            .Where(i => i.DisplayEnd > i.DisplayStart)
            .OrderBy(i => i.DisplayStart)
            .ThenBy(i => i.DisplayEnd)
            .ToList();

        var group = new List<WeekCalendarItem>();
        var groupEnd = DateTime.MinValue;

        foreach (var item in visible)
        {
            if (group.Count == 0)
            {
                group.Add(item);
                groupEnd = item.DisplayEnd;
                continue;
            }

            if (item.DisplayStart < groupEnd)
            {
                group.Add(item);
                if (item.DisplayEnd > groupEnd) groupEnd = item.DisplayEnd;
                continue;
            }

            AssignOverlapLayout(group, rangeStart);
            group.Clear();
            group.Add(item);
            groupEnd = item.DisplayEnd;
        }

        if (group.Count > 0)
            AssignOverlapLayout(group, rangeStart);

        foreach (var item in items.Where(i => i.DisplayEnd <= i.DisplayStart))
        {
            item.DisplayWidth = 0;
            item.DisplayHeight = 0;
        }
    }

    private void AssignOverlapLayout(List<WeekCalendarItem> group, DateTime rangeStart)
    {
        var columnsEnd = new List<DateTime>();

        foreach (var item in group.OrderBy(i => i.DisplayStart).ThenBy(i => i.DisplayEnd))
        {
            var placed = false;
            for (var col = 0; col < columnsEnd.Count; col++)
            {
                if (item.DisplayStart >= columnsEnd[col])
                {
                    item.OverlapColumn = col;
                    columnsEnd[col] = item.DisplayEnd;
                    placed = true;
                    break;
                }
            }

            if (!placed)
            {
                item.OverlapColumn = columnsEnd.Count;
                columnsEnd.Add(item.DisplayEnd);
            }
        }

        var columnCount = Math.Max(1, columnsEnd.Count);
        var availableWidth = DayColumnWidth - (DayInnerPadding * 2);
        var blockWidth = Math.Max(46, (availableWidth - ((columnCount - 1) * OverlapGap)) / columnCount);

        foreach (var item in group)
        {
            item.OverlapColumnCount = columnCount;
            item.DisplayLeft = DayInnerPadding + (item.OverlapColumn * (blockWidth + OverlapGap));
            item.DisplayWidth = blockWidth;
            item.DisplayTop = (item.DisplayStart - rangeStart).TotalMinutes * PixelsPerMinute;
            item.DisplayHeight = Math.Max(28, (item.DisplayEnd - item.DisplayStart).TotalMinutes * PixelsPerMinute - 2);
            item.TimeLabel = $"{item.SegmentStart:HH:mm} - {item.SegmentEnd:HH:mm}";
            item.IsCompact = item.DisplayHeight < 46;
            item.ShowNote = item.DisplayHeight >= 64 && !string.IsNullOrWhiteSpace(item.SegmentNote);
            item.ShowTime = item.DisplayHeight >= 34;
        }

        // Prevent visual overlap caused by enforced minimum height.
        // We only adjust vertical rendering position, not the actual segment time data.
        foreach (var colGroup in group.GroupBy(g => g.OverlapColumn))
        {
            var colItems = colGroup.OrderBy(i => i.DisplayTop).ToList();
            for (var i = 1; i < colItems.Count; i++)
            {
                var prev = colItems[i - 1];
                var current = colItems[i];
                var minTop = prev.DisplayTop + prev.DisplayHeight + 2;
                if (current.DisplayTop < minTop)
                    current.DisplayTop = minTop;
            }
        }
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

    private static DateTime StartOfWeek(DateTime date)
    {
        var firstDay = CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek;
        var diff = (7 + (date.DayOfWeek - firstDay)) % 7;
        return date.Date.AddDays(-diff);
    }

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

    private bool _isToday;
    public bool IsToday { get => _isToday; set => Set(ref _isToday, value); }
}

public class WeekCalendarItem
{
    public Guid TaskId { get; set; }
    public string TaskTitle { get; set; } = string.Empty;
    public string TaskDescription { get; set; } = string.Empty;
    public string TicketUrl { get; set; } = string.Empty;
    public long SegmentId { get; set; }
    public DateTime SegmentStart { get; set; }
    public DateTime SegmentEnd { get; set; }
    public int SegmentIndexDisplay { get; set; }
    public string TaskStatus { get; set; } = string.Empty;
    public string SegmentNote { get; set; } = string.Empty;

    public DateTime DisplayStart { get; set; }
    public DateTime DisplayEnd { get; set; }
    public double DisplayTop { get; set; }
    public double DisplayHeight { get; set; }
    public double DisplayLeft { get; set; }
    public double DisplayWidth { get; set; }
    public int OverlapColumn { get; set; }
    public int OverlapColumnCount { get; set; }
    public string TimeLabel { get; set; } = string.Empty;
    public bool IsCompact { get; set; }
    public bool ShowNote { get; set; }
    public bool ShowTime { get; set; }

    public string TooltipText =>
        $"{TaskTitle}\n{TimeLabel}\nStatus: {TaskStatus}" +
        (string.IsNullOrWhiteSpace(SegmentNote) ? string.Empty : $"\nNotiz: {SegmentNote}") +
        (string.IsNullOrWhiteSpace(TaskDescription) ? string.Empty : $"\n{TaskDescription}") +
        (string.IsNullOrWhiteSpace(TicketUrl) ? string.Empty : $"\nTicket: {TicketUrl}");
}

public class TimeAxisLabel
{
    public string Label { get; set; } = string.Empty;
    public double Top { get; set; }
}

public class TimeGridLine
{
    public double Top { get; set; }
    public bool IsHour { get; set; }
}
