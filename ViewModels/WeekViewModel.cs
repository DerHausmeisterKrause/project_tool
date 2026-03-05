using System.Collections.Generic;
using System;
using System.Linq;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows;
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
    private readonly OutlookCalendarService _outlookCalendar;

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
                SelectedDay = _selectedDate.Date;
                var day = Days.FirstOrDefault(d => d.DayDate.Date == _selectedDate.Date);
                if (day != null && !ReferenceEquals(day, SelectedDayGroup))
                    SelectedDayGroup = day;
            }
        }
    }

    private DateTime _selectedDay = DateTime.Today.Date;
    public DateTime SelectedDay
    {
        get => _selectedDay;
        set => Set(ref _selectedDay, value.Date);
    }

    private WeekDayGroup? _selectedDayGroup;
    public WeekDayGroup? SelectedDayGroup
    {
        get => _selectedDayGroup;
        set
        {
            if (Set(ref _selectedDayGroup, value))
            {
                if (value != null && value.DayDate.Date != SelectedDate.Date)
                    Set(ref _selectedDate, value.DayDate.Date, nameof(SelectedDate));

                if (value != null)
                    Set(ref _selectedDay, value.DayDate.Date, nameof(SelectedDay));

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

    public string SelectedDayType => SelectedDayGroup?.DayType ?? "Normal";
    public bool SelectedIsHo => SelectedDayGroup?.IsHo ?? false;
    public bool SelectedIsBr => SelectedDayGroup?.IsBr ?? false;

    public RelayCommand PreviousWeekCommand { get; }
    public RelayCommand NextWeekCommand { get; }
    public RelayCommand CurrentWeekCommand { get; }
    public RelayCommand<WeekDayGroup> SelectDayCommand { get; }
    public RelayCommand<WeekCalendarItem> OpenCalendarItemCommand { get; }
    public RelayCommand<string> OpenTicketUrlCommand { get; }
    public RelayCommand<string> OpenTeamsUrlCommand { get; }
    public RelayCommand<OutlookCalendarBlock> OpenOutlookEventCommand { get; }
    public RelayCommand SetDayTypeNormalCommand { get; }
    public RelayCommand SetDayTypeUlCommand { get; }
    public RelayCommand SetDayTypeAmCommand { get; }
    public RelayCommand ToggleHoCommand { get; }
    public RelayCommand ToggleBrCommand { get; }

    public WeekViewModel(TaskService tasks, WorkDayService workDays, SettingsService settings, OutlookCalendarService outlookCalendar)
    {
        _tasks = tasks;
        _workDays = workDays;
        _settings = settings;
        _outlookCalendar = outlookCalendar;

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
        OpenTeamsUrlCommand = new RelayCommand<string>(OpenTeamsUrlFromWeek, url => !string.IsNullOrWhiteSpace(url));
        OpenOutlookEventCommand = new RelayCommand<OutlookCalendarBlock>(OpenOutlookEvent, evt => evt != null);
        SetDayTypeNormalCommand = new RelayCommand(() => SetDayType("Normal"), () => SelectedDayGroup != null);
        SetDayTypeUlCommand = new RelayCommand(() => SetDayType("UL"), () => SelectedDayGroup != null);
        SetDayTypeAmCommand = new RelayCommand(() => SetDayType("AM"), () => SelectedDayGroup != null);
        ToggleHoCommand = new RelayCommand(() => { if (SelectedDayGroup == null) return; SelectedDayGroup.IsHo = !SelectedDayGroup.IsHo; SaveSelectedDay(); }, () => SelectedDayGroup != null);
        ToggleBrCommand = new RelayCommand(() => { if (SelectedDayGroup == null) return; SelectedDayGroup.IsBr = !SelectedDayGroup.IsBr; SaveSelectedDay(); }, () => SelectedDayGroup != null);

        _nowIndicatorTimer.Interval = TimeSpan.FromSeconds(30);
        _nowIndicatorTimer.Tick += (_, _) => UpdateNowIndicator();
        _nowIndicatorTimer.Start();

        _outlookCalendar.EventsUpdated += OnOutlookEventsUpdated;
        LoadWeek();
        _ = _outlookCalendar.TriggerSyncAsync("week-init");
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
        ServiceLocator.MainViewModel.NavigateToTodayAndOpenTask(item.TaskId);
    }

    private void OpenTicketUrlFromWeek(string? url)
    {
        UrlLauncher.TryOpen(url, out _);
    }

    private void OpenTeamsUrlFromWeek(string? url)
    {
        UrlLauncher.TryOpen(url, out _);
    }

    private void OpenOutlookEvent(OutlookCalendarBlock? evt)
    {
        if (evt == null)
            return;

        var opened = ServiceLocator.Outlook.OpenCalendarEvent(evt.Id);
        if (opened.ok)
            return;

        if (!string.IsNullOrWhiteSpace(evt.TeamsJoinUrl))
        {
            UrlLauncher.TryOpen(evt.TeamsJoinUrl, out _);
            return;
        }

        var details = $"Outlook-Termin konnte nicht geöffnet werden.

{opened.error}

{evt.Subject}
{evt.TimeLabel}";
        MessageBox.Show(details, "Outlook", MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void OnOutlookEventsUpdated()
    {
        App.Current?.Dispatcher.Invoke(LoadWeek);
    }

    private void SetDayType(string type)
    {
        if (SelectedDayGroup == null) return;
        SelectedDayGroup.DayType = type;
        SaveSelectedDay();
    }

    private void SaveSelectedDay()
    {
        if (SelectedDayGroup == null) return;
        _workDays.SetDayMarkers(SelectedDayGroup.DayDate.ToString("yyyy-MM-dd"), SelectedDayGroup.DayType, SelectedDayGroup.IsBr, SelectedDayGroup.IsHo);
        var selectedDate = SelectedDayGroup.DayDate;
        LoadWeek();
        SelectedDayGroup = Days.FirstOrDefault(d => d.DayDate.Date == selectedDate.Date) ?? Days.FirstOrDefault();
    }

    private void LoadWeek()
    {
        var previousSelectionDate = SelectedDayGroup?.DayDate;

        if (_settings.Current.OutlookCalendarEnabled)
            _ = _outlookCalendar.TriggerSyncAsync("week-load");

        Days.Clear();
        var from = WeekStart.Date;
        var to = WeekStart.AddDays(7).Date;
        var outlookEvents = _outlookCalendar.GetEvents(from, to);

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

            LayoutDayItemsCore(day.Date, calendarItems);

            var markerResult = ResolveOutlookDerivedMarker(day.Date, outlookEvents);
            var external = BuildExternalEventsForDay(day.Date, outlookEvents, markerResult.ConsumedEventIds);
            ApplySharedOverlapLayout(calendarItems, external);
            MarkSegmentConflicts(calendarItems, external);

            var displayDayType = wd.DayType;
            if (displayDayType == "Normal" && (markerResult.DerivedDayType == "UL" || markerResult.DerivedDayType == "AM"))
                displayDayType = markerResult.DerivedDayType;

            var displayIsHo = wd.IsHo || markerResult.DerivedHo;

            Days.Add(new WeekDayGroup
            {
                DayDate = day,
                DayLabel = day.ToString("ddd dd.MM", CultureInfo.CurrentCulture),
                IsToday = day.Date == DateTime.Today,
                CalendarItems = new ObservableCollection<WeekCalendarItem>(calendarItems.OrderBy(c => c.DisplayTop)),
                ExternalEvents = new ObservableCollection<OutlookCalendarBlock>(external),
                DayType = displayDayType,
                IsBr = wd.IsBr,
                IsHo = displayIsHo,
                Summary = $"Soll {Fmt(target)} | Ist {Fmt(net)} | Ü {Fmt(overtime)}"
            });
        }

        SelectedDayGroup = ResolveSelectedDay(previousSelectionDate);
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
        NowLineTop = MapToCalendarY(now, today);
        NowMarkerLeft = dayIndex * DayColumnWidth - 4;
        NowMarkerTop = NowLineTop - 4;
        ShowNowIndicator = true;
    }

    private double MapToCalendarY(DateTime value, DateTime dayDate)
    {
        var dayStart = dayDate.Date.AddHours(CalendarStartHour);
        return (value - dayStart).TotalMinutes * PixelsPerMinute;
    }

    private void LayoutDayItemsCore(DateTime dayDate, List<WeekCalendarItem> items)
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

            AssignOverlapLayoutCore(group, rangeStart);
            group.Clear();
            group.Add(item);
            groupEnd = item.DisplayEnd;
        }

        if (group.Count > 0)
            AssignOverlapLayoutCore(group, rangeStart);

        foreach (var item in items.Where(i => i.DisplayEnd <= i.DisplayStart))
        {
            item.DisplayWidth = 0;
            item.DisplayHeight = 0;
        }
    }

    private void AssignOverlapLayoutCore(List<WeekCalendarItem> group, DateTime rangeStart)
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
            item.DisplayTop = MapToCalendarY(item.DisplayStart, dayDate: rangeStart.Date);
            item.DisplayHeight = Math.Max(28, (item.DisplayEnd - item.DisplayStart).TotalMinutes * PixelsPerMinute - 2);
            item.TimeLabel = $"{item.SegmentStart:HH:mm} - {item.SegmentEnd:HH:mm}";
            var durationMinutes = (item.DisplayEnd - item.DisplayStart).TotalMinutes;
            item.IsCompact = durationMinutes < 30 || item.DisplayHeight < 40;
            item.ShowNote = !item.IsCompact && item.DisplayHeight >= 64 && !string.IsNullOrWhiteSpace(item.SegmentNote);
            item.ShowTime = item.IsCompact || item.DisplayHeight >= 34;
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

    private (string DerivedDayType, bool DerivedHo, HashSet<string> ConsumedEventIds) ResolveOutlookDerivedMarker(DateTime dayDate, IReadOnlyList<OutlookCalendarEvent> events)
    {
        var consumed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (!_settings.Current.OutlookInterpretAllDayAsMarkers)
            return ("Normal", false, consumed);

        var dayStart = dayDate.Date;
        var dayEnd = dayDate.Date.AddDays(1);

        string derivedDayType = "Normal";
        var derivedHo = false;

        foreach (var evt in events.Where(e => e.EndLocal > dayStart && e.StartLocal < dayEnd))
        {
            var duration = evt.EndLocal - evt.StartLocal;
            var eligible = evt.IsAllDay || duration.TotalHours >= 6;
            if (!eligible)
                continue;

            if (TryMapDayMarker(evt.Subject, out var mapped))
            {
                consumed.Add(evt.Id);
                if (mapped == "HO")
                    derivedHo = true;
                else
                    derivedDayType = mapped;
            }
        }

        return (derivedDayType, derivedHo, consumed);
    }

    private static bool TryMapDayMarker(string subject, out string mapped)
    {
        mapped = "Normal";
        if (string.IsNullOrWhiteSpace(subject))
            return false;

        var s = subject.ToLowerInvariant();
        if (s.Contains("homeoffice") || s.Contains(" ho ") || s.StartsWith("ho") || s.EndsWith("ho"))
        {
            mapped = "HO";
            return true;
        }

        if (s.Contains("urlaub") || s.Contains(" ul ") || s.StartsWith("ul") || s.EndsWith("ul"))
        {
            mapped = "UL";
            return true;
        }

        if (s.Contains("maz"))
        {
            mapped = "AM";
            return true;
        }

        return false;
    }

    private void ApplySharedOverlapLayout(List<WeekCalendarItem> segments, List<OutlookCalendarBlock> external)
    {
        var blocks = new List<LayoutBlockRef>();
        blocks.AddRange(segments.Where(s => s.DisplayEnd > s.DisplayStart).Select(s => new LayoutBlockRef(s.DisplayStart, s.DisplayEnd,
            (col, count) =>
            {
                s.OverlapColumn = col;
                s.OverlapColumnCount = count;
            })));

        blocks.AddRange(external.Where(e => e.EndLocal > e.StartLocal).Select(e => new LayoutBlockRef(e.StartLocal, e.EndLocal,
            (col, count) =>
            {
                e.OverlapColumn = col;
                e.OverlapColumnCount = count;
            })));

        if (blocks.Count == 0)
            return;

        var sorted = blocks.OrderBy(b => b.Start).ThenBy(b => b.End).ToList();
        var group = new List<LayoutBlockRef>();
        var groupEnd = DateTime.MinValue;

        foreach (var block in sorted)
        {
            if (group.Count == 0)
            {
                group.Add(block);
                groupEnd = block.End;
                continue;
            }

            if (block.Start < groupEnd)
            {
                group.Add(block);
                if (block.End > groupEnd)
                    groupEnd = block.End;
                continue;
            }

            AssignSharedGroup(group, segments, external);
            group.Clear();
            group.Add(block);
            groupEnd = block.End;
        }

        if (group.Count > 0)
            AssignSharedGroup(group, segments, external);
    }

    private void AssignSharedGroup(List<LayoutBlockRef> group, List<WeekCalendarItem> segments, List<OutlookCalendarBlock> external)
    {
        var columnsEnd = new List<DateTime>();
        foreach (var block in group.OrderBy(i => i.Start).ThenBy(i => i.End))
        {
            var placed = false;
            for (var col = 0; col < columnsEnd.Count; col++)
            {
                if (block.Start >= columnsEnd[col])
                {
                    block.Column = col;
                    columnsEnd[col] = block.End;
                    placed = true;
                    break;
                }
            }

            if (!placed)
            {
                block.Column = columnsEnd.Count;
                columnsEnd.Add(block.End);
            }
        }

        var columnCount = Math.Max(1, columnsEnd.Count);
        var availableWidth = DayColumnWidth - (DayInnerPadding * 2);
        var blockWidth = Math.Max(42, (availableWidth - ((columnCount - 1) * OverlapGap)) / columnCount);

        foreach (var block in group)
            block.Assign(block.Column, columnCount);

        foreach (var seg in segments.Where(s => s.OverlapColumnCount == columnCount && s.DisplayEnd > s.DisplayStart && group.Any(g => g.Start == s.DisplayStart && g.End == s.DisplayEnd)))
        {
            seg.DisplayLeft = DayInnerPadding + (seg.OverlapColumn * (blockWidth + OverlapGap));
            seg.DisplayWidth = blockWidth;
        }

        foreach (var ext in external.Where(e => e.OverlapColumnCount == columnCount && group.Any(g => g.Start == e.StartLocal && g.End == e.EndLocal)))
        {
            ext.DisplayLeft = DayInnerPadding + (ext.OverlapColumn * (blockWidth + OverlapGap));
            ext.DisplayWidth = blockWidth;
        }
    }

    private List<OutlookCalendarBlock> BuildExternalEventsForDay(DateTime dayDate, IReadOnlyList<OutlookCalendarEvent> source, HashSet<string> consumedEventIds)
    {
        var dayStart = dayDate.Date.AddHours(CalendarStartHour);
        var dayEnd = dayDate.Date.AddHours(CalendarEndHour);

        return source
             .Where(e => e.EndLocal > dayStart && e.StartLocal < dayEnd && !consumedEventIds.Contains(e.Id))
            .Select(e =>
            {
                var start = e.StartLocal < dayStart ? dayStart : e.StartLocal;
                var end = e.EndLocal > dayEnd ? dayEnd : e.EndLocal;
                var top = MapToCalendarY(start, dayDate);
                var height = Math.Max(24, (end - start).TotalMinutes * PixelsPerMinute - 2);
                return new OutlookCalendarBlock
                {
                    Id = e.Id,
                    StartLocal = e.StartLocal,
                    EndLocal = e.EndLocal,
                    Subject = e.Subject,
                    TimeLabel = $"{e.StartLocal:HH:mm} - {e.EndLocal:HH:mm}",
                    Location = e.Location,
                    TeamsJoinUrl = _settings.Current.OutlookTeamsButtonEnabled ? e.OnlineMeetingJoinUrl : string.Empty,
                    DisplayTop = top,
                    DisplayHeight = height,
                    DisplayLeft = DayInnerPadding,
                    DisplayWidth = Math.Max(46, DayColumnWidth - (DayInnerPadding * 2)),
                    TooltipText = $"Outlook: {e.Subject}\n{e.StartLocal:HH:mm} - {e.EndLocal:HH:mm}" +
                                  (string.IsNullOrWhiteSpace(e.Location) ? string.Empty : $"\nOrt: {e.Location}")
                };
            })
            .OrderBy(e => e.DisplayTop)
            .ToList();
    }

    private static void MarkSegmentConflicts(List<WeekCalendarItem> segments, List<OutlookCalendarBlock> external)
    {
        foreach (var segment in segments)
        {
            var conflict = external.FirstOrDefault(e => segment.SegmentEnd > e.StartLocal && segment.SegmentStart < e.EndLocal);
            segment.HasOutlookConflict = conflict != null;
            segment.OutlookConflictText = conflict == null ? string.Empty : $"Konflikt mit Outlook-Termin: {conflict.Subject}";
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
            return Days.FirstOrDefault(d => d.DayDate.Date == SelectedDay.Date)
               ?? Days.FirstOrDefault(d => d.DayDate.Date == today.Date)
               ?? Days.First();

        return Days.First();
    }

    private static string Fmt(int minutes) => $"{minutes / 60}h {Math.Abs(minutes % 60):00}m";

    private static DateTime StartOfWeek(DateTime date)
    {
        var firstDay = CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek;
        var diff = (7 + (date.DayOfWeek - firstDay)) % 7;
        return date.Date.AddDays(-diff);
    }

    public void Refresh()
    {
        _ = _outlookCalendar.TriggerSyncAsync("week-refresh");
        LoadWeek();
    }

    public override string ToString() => Title;
}

public class WeekDayGroup : ObservableObject
{
    public DateTime DayDate { get; set; }
    public string DayLabel { get; set; } = string.Empty;
    public ObservableCollection<WeekCalendarItem> CalendarItems { get; set; } = new();
    public ObservableCollection<OutlookCalendarBlock> ExternalEvents { get; set; } = new();

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
    public bool HasOutlookConflict { get; set; }
    public string OutlookConflictText { get; set; } = string.Empty;

    public string TooltipText =>
        $"{TaskTitle}\n{TimeLabel}\nStatus: {TaskStatus}" +
        (string.IsNullOrWhiteSpace(SegmentNote) ? string.Empty : $"\nNotiz: {SegmentNote}") +
        (string.IsNullOrWhiteSpace(TaskDescription) ? string.Empty : $"\n{TaskDescription}") +
        (string.IsNullOrWhiteSpace(TicketUrl) ? string.Empty : $"\nTicket: {TicketUrl}") +
        (string.IsNullOrWhiteSpace(OutlookConflictText) ? string.Empty : $"\n{OutlookConflictText}");
}

public class OutlookCalendarBlock
{
    public string Id { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public string TimeLabel { get; set; } = string.Empty;
    public string Location { get; set; } = string.Empty;
    public string TeamsJoinUrl { get; set; } = string.Empty;
    public DateTime StartLocal { get; set; }
    public DateTime EndLocal { get; set; }
    public double DisplayTop { get; set; }
    public double DisplayHeight { get; set; }
    public double DisplayLeft { get; set; }
    public double DisplayWidth { get; set; }
    public int OverlapColumn { get; set; }
    public int OverlapColumnCount { get; set; }
    public bool HasTeamsLink => !string.IsNullOrWhiteSpace(TeamsJoinUrl);
    public string TooltipText { get; set; } = string.Empty;
}



internal sealed class LayoutBlockRef
{
    public DateTime Start { get; }
    public DateTime End { get; }
    public Action<int, int> Assign { get; }
    public int Column { get; set; }

    public LayoutBlockRef(DateTime start, DateTime end, Action<int, int> assign)
    {
        Start = start;
        End = end;
        Assign = assign;
    }
}

public class OutlookCalendarBlock
{
    public string Id { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public string TimeLabel { get; set; } = string.Empty;
    public string Location { get; set; } = string.Empty;
    public string TeamsJoinUrl { get; set; } = string.Empty;
    public DateTime StartLocal { get; set; }
    public DateTime EndLocal { get; set; }
    public double DisplayTop { get; set; }
    public double DisplayHeight { get; set; }
    public double DisplayLeft { get; set; }
    public double DisplayWidth { get; set; }
    public bool HasTeamsLink => !string.IsNullOrWhiteSpace(TeamsJoinUrl);
    public string TooltipText { get; set; } = string.Empty;
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
