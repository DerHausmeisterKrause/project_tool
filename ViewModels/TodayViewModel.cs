using System.Linq;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Threading;
using TaskTool.Infrastructure;
using TaskTool.Models;
using TaskStatus = TaskTool.Models.TaskStatus;
using TaskTool.Services;

namespace TaskTool.ViewModels;

public class TodayViewModel : ObservableObject
{
    private readonly TaskService _tasks;
    private readonly WorkDayService _workDays;
    private readonly SettingsService _settings;
    private readonly GermanTimeService _germanTime;
    private readonly DispatcherTimer _clock;
    private readonly OutlookCalendarService _outlookCalendar;
    private readonly TicketSystemService _ticketSystem;
    private DateTime _agendaDate;

    public string Title => "Heute";
    // Existing timer and Dynamic Island consumers use this as the complete active-task collection.
    public ObservableCollection<TaskItem> CurrentTasks { get; } = new();
    public ObservableCollection<TaskItem> TodayTasks { get; } = new();
    private readonly ObservableCollection<TaskItem> _currentTasksWithoutToday = new();
    public ObservableCollection<TaskItem> DisplayedTasks { get; } = new();
    public ObservableCollection<TaskItem> CompletedTasks { get; } = new();
    public ObservableCollection<TodayAgendaItem> TodayAgendaItems { get; } = new();
    public ObservableCollection<BreakEditRow> BreakRows { get; } = new();
    public ObservableCollection<TaskSegment> Segments { get; } = new();
    public ObservableCollection<TicketTimeBooking> TicketTimeBookings { get; } = new();
    public ObservableCollection<TicketFieldOption> CostCenterOptions { get; } = new();
    public ObservableCollection<TicketFieldOption> OrderOptions { get; } = new();
    public ObservableCollection<string> TimeOptions { get; } = new(Enumerable.Range(0, 96).Select(i => TimeSpan.FromMinutes(i * 15).ToString(@"hh\:mm")));
    public IReadOnlyList<string> CurrentTaskSortFields { get; } = new[] { "Zuletzt bearbeitet", "Erstellungsdatum" };
    public IReadOnlyList<string> CurrentTaskSortDirections { get; } = new[] { "Neueste zuerst", "Älteste zuerst" };

    public event Action<Guid>? TaskBringIntoViewRequested;

    private TaskItem? _selectedTask;
    public TaskItem? SelectedTask
    {
        get => _selectedTask;
        set
        {
            if (Set(ref _selectedTask, value))
            {
                LoadSegments();
                Raise(nameof(IsTaskSelected));
                Raise(nameof(HasZnunyTicket));
                RaiseCommandStates();
                UpdateTimerDisplay();
                LoadTicketBookingHistory();
                _ = LoadTicketBookingContextAsync(value);
            }
        }
    }

    public bool IsTaskSelected => SelectedTask != null;

    private string _quickAddText = string.Empty;
    public string QuickAddText { get => _quickAddText; set => Set(ref _quickAddText, value); }

    private string _taskSearchText = string.Empty;
    public string TaskSearchText
    {
        get => _taskSearchText;
        set { if (Set(ref _taskSearchText, value)) ApplyTaskFilters(); }
    }

    private string _completedTaskSearchText = string.Empty;
    public string CompletedTaskSearchText
    {
        get => _completedTaskSearchText;
        set { if (Set(ref _completedTaskSearchText, value)) ApplyTaskFilters(); }
    }

    private TodayTaskScope _selectedTaskScope = TodayTaskScope.Today;
    public TodayTaskScope SelectedTaskScope
    {
        get => _selectedTaskScope;
        set
        {
            if (Set(ref _selectedTaskScope, value))
            {
                RefreshDisplayedTasks();
                Raise(nameof(ShowActiveTaskList));
                Raise(nameof(ShowTodayAgenda));
                Raise(nameof(ShowCurrentTaskList));
                Raise(nameof(ShowCompletedTaskList));
                Raise(nameof(ActiveTaskListHeading));
            }
        }
    }

    public bool ShowActiveTaskList => SelectedTaskScope != TodayTaskScope.Completed;
    public bool ShowTodayAgenda => SelectedTaskScope == TodayTaskScope.Today;
    public bool ShowCurrentTaskList => SelectedTaskScope == TodayTaskScope.Current;
    public bool ShowCompletedTaskList => SelectedTaskScope == TodayTaskScope.Completed;
    public string ActiveTaskListHeading => SelectedTaskScope == TodayTaskScope.Today ? "Heute:" : "Aktuelle Aufgaben:";

    public string SelectedCurrentTaskSortField
    {
        get => string.Equals(_settings.Current.CurrentTasksSortField, "Created", StringComparison.OrdinalIgnoreCase)
            ? "Erstellungsdatum"
            : "Zuletzt bearbeitet";
        set
        {
            var field = string.Equals(value, "Erstellungsdatum", StringComparison.Ordinal) ? "Created" : "Updated";
            if (string.Equals(_settings.Current.CurrentTasksSortField, field, StringComparison.Ordinal)) return;
            _settings.Current.CurrentTasksSortField = field;
            _settings.Save();
            Raise();
            RefreshDisplayedTasks();
        }
    }

    public string SelectedCurrentTaskSortDirection
    {
        get => _settings.Current.CurrentTasksSortDescending ? "Neueste zuerst" : "Älteste zuerst";
        set
        {
            var descending = !string.Equals(value, "Älteste zuerst", StringComparison.Ordinal);
            if (_settings.Current.CurrentTasksSortDescending == descending) return;
            _settings.Current.CurrentTasksSortDescending = descending;
            _settings.Save();
            Raise();
            RefreshDisplayedTasks();
        }
    }

    private string _statusMessage = string.Empty;
    public string StatusMessage { get => _statusMessage; set => Set(ref _statusMessage, value); }

    private string _workDaySummary = string.Empty;
    public string WorkDaySummary { get => _workDaySummary; set => Set(ref _workDaySummary, value); }

    private int _ticketMinutesToday;
    public int TicketMinutesToday { get => _ticketMinutesToday; set => Set(ref _ticketMinutesToday, value); }

    private int _ticketMinutesCurrentMonth;
    public int TicketMinutesCurrentMonth { get => _ticketMinutesCurrentMonth; set => Set(ref _ticketMinutesCurrentMonth, value); }

    private string _comeTimeText = string.Empty;
    public string ComeTimeText { get => _comeTimeText; set => Set(ref _comeTimeText, value); }

    private string _goTimeText = string.Empty;
    public string GoTimeText { get => _goTimeText; set => Set(ref _goTimeText, value); }

    private string _timerDisplay = "00:00:00";
    public string TimerDisplay { get => _timerDisplay; set => Set(ref _timerDisplay, value); }

    private TicketFieldOption? _selectedCostCenter;
    public TicketFieldOption? SelectedCostCenter { get => _selectedCostCenter; set => Set(ref _selectedCostCenter, value); }

    private TicketFieldOption? _selectedOrder;
    public TicketFieldOption? SelectedOrder { get => _selectedOrder; set => Set(ref _selectedOrder, value); }

    private string _ticketBookingInformation = string.Empty;
    public string TicketBookingInformation { get => _ticketBookingInformation; set => Set(ref _ticketBookingInformation, value); }

    private bool _isTicketBooking;
    public bool IsTicketBooking
    {
        get => _isTicketBooking;
        set
        {
            if (Set(ref _isTicketBooking, value))
            {
                BookTimeInTicketSystemCommand.RaiseCanExecuteChanged();
                RefreshTicketFieldOptionsCommand.RaiseCanExecuteChanged();
                CheckTicketTimeBookingCommand.RaiseCanExecuteChanged();
                RetryTicketTimeBookingCommand.RaiseCanExecuteChanged();
            }
        }
    }

    private long _successfullyTransferredSeconds;
    private long _ticketTimeBookingBaselineSeconds;
    private decimal _successfullyBookedMinutes;
    private bool _hasUnresolvedTicketTimeBooking;
    public long UnbookedTicketSeconds => Math.Max(0, (SelectedTask?.TicketSecondsBooked ?? 0) + (long)(SelectedTask == null ? TimeSpan.Zero : _tasks.GetOpenSessionDuration(SelectedTask.Id)).TotalSeconds - _ticketTimeBookingBaselineSeconds - _successfullyTransferredSeconds);
    public long SuccessfullyTransferredSeconds => _successfullyTransferredSeconds;
    public string UnbookedTicketTimeText => $"Noch nicht gebucht: {UnbookedTicketSeconds / 60m:0.##} Min.";
    public string TransferredTicketTimeText => $"Insgesamt über TaskTool in OTRS gebucht: {_successfullyBookedMinutes:0.##} Min.";
    public bool HasZnunyTicket => SelectedTask?.Tags.Contains("ZnunyTicketID:", StringComparison.OrdinalIgnoreCase) == true;

    private DateTime? _newSegmentDate = DateTime.Today;
    public DateTime? NewSegmentDate
    {
        get => _newSegmentDate;
        set { if (Set(ref _newSegmentDate, value)) RaiseSegmentEditorState(); }
    }

    private string _newSegmentStartTime = "09:00";
    public string NewSegmentStartTime
    {
        get => _newSegmentStartTime;
        set
        {
            if (Set(ref _newSegmentStartTime, value))
            {
                if (string.IsNullOrWhiteSpace(NewSegmentEndTime) && TimeSpan.TryParse(value, out var start))
                    NewSegmentEndTime = start.Add(TimeSpan.FromMinutes(30)).ToString(@"hh\:mm");
                RaiseSegmentEditorState();
            }
        }
    }

    private string _newSegmentEndTime = "09:30";
    public string NewSegmentEndTime
    {
        get => _newSegmentEndTime;
        set { if (Set(ref _newSegmentEndTime, value)) RaiseSegmentEditorState(); }
    }

    private string _newSegmentNote = string.Empty;
    public string NewSegmentNote { get => _newSegmentNote; set => Set(ref _newSegmentNote, value); }

    private string _newSegmentConflictWarning = string.Empty;
    public string NewSegmentConflictWarning
    {
        get => _newSegmentConflictWarning;
        set
        {
            if (Set(ref _newSegmentConflictWarning, value))
                Raise(nameof(HasNewSegmentConflict));
        }
    }

    public bool HasNewSegmentConflict => !string.IsNullOrWhiteSpace(NewSegmentConflictWarning);

    public string NewSegmentValidationHint
    {
        get
        {
            if (NewSegmentDate == null) return "Datum muss gesetzt sein.";
            if (!TimeSpan.TryParse(NewSegmentStartTime, out var start)) return "Startzeit ungültig (HH:mm).";
            if (string.IsNullOrWhiteSpace(NewSegmentEndTime)) return "Endzeit darf nicht leer sein.";
            if (!TimeSpan.TryParse(NewSegmentEndTime, out var end)) return "Endzeit ungültig (HH:mm).";
            if (start >= end) return "Startzeit muss vor Endzeit liegen.";
            return string.Empty;
        }
    }

    public bool CanSaveNewSegment => SelectedTask != null && string.IsNullOrWhiteSpace(NewSegmentValidationHint);

    public RelayCommand QuickAddCommand { get; }
    public RelayCommand SaveCommand { get; }
    public RelayCommand ReopenCommand { get; }
    public RelayCommand DoneCommand { get; }
    public RelayCommand StartTimerCommand { get; }
    public RelayCommand StopTimerCommand { get; }
    public RelayCommand Add15Command { get; }
    public RelayCommand Add30Command { get; }
    public RelayCommand Add60Command { get; }
    public RelayCommand Subtract15Command { get; }
    public RelayCommand Subtract30Command { get; }
    public RelayCommand Subtract60Command { get; }
    public RelayCommand BookTimeInTicketSystemCommand { get; }
    public RelayCommand RefreshTicketFieldOptionsCommand { get; }
    public RelayCommand<TicketTimeBooking> CheckTicketTimeBookingCommand { get; }
    public RelayCommand<TicketTimeBooking> RetryTicketTimeBookingCommand { get; }
    public RelayCommand ComeCommand { get; }
    public RelayCommand GoCommand { get; }
    public RelayCommand BreakStartCommand { get; }
    public RelayCommand BreakEndCommand { get; }
    public RelayCommand ManualSaveCommand { get; }
    public RelayCommand AddBreakRowCommand { get; }
    public RelayCommand SaveMarkersCommand { get; }
    public RelayCommand SetDayTypeNormalCommand { get; }
    public RelayCommand SetDayTypeAmCommand { get; }
    public RelayCommand SetDayTypeUlCommand { get; }
    public RelayCommand AddSegmentCommand { get; }
    public RelayCommand ShowTodayTasksCommand { get; }
    public RelayCommand ShowCurrentTasksCommand { get; }
    public RelayCommand ShowCompletedTasksCommand { get; }

    public RelayCommand<TaskItem> SelectTaskCommand { get; }
    public RelayCommand<TaskItem> StartTaskCommand { get; }
    public RelayCommand<TaskItem> StopTaskCommand { get; }
    public RelayCommand<TaskItem> DoneTaskCommand { get; }
    public RelayCommand<TaskItem> TogglePinTaskCommand { get; }
    public RelayCommand<string> OpenTicketUrlCommand { get; }
    public RelayCommand<OutlookCalendarEvent> OpenAgendaOutlookEventCommand { get; }
    public RelayCommand<string> OpenAgendaTeamsCommand { get; }
    public RelayCommand<TaskSegment> SaveSegmentCommand { get; }
    public RelayCommand<TaskSegment> DeleteSegmentCommand { get; }
    public RelayCommand<TaskSegment> DeleteSegmentOutlookCommand { get; }

    private string _dayType = "Normal";
    public string DayType { get => _dayType; set => Set(ref _dayType, value); }

    private bool _isBr;
    public bool IsBr { get => _isBr; set => Set(ref _isBr, value); }

    private bool _isHo;
    public bool IsHo { get => _isHo; set => Set(ref _isHo, value); }

    public TodayViewModel(TaskService tasks, WorkDayService workDays, SettingsService settings, OutlookCalendarService outlookCalendar, TicketSystemService ticketSystem)
    {
        _tasks = tasks;
        _workDays = workDays;
        _settings = settings;
        _germanTime = ServiceLocator.GermanTime;
        _outlookCalendar = outlookCalendar;
        _ticketSystem = ticketSystem;
        _tasks.SegmentsChanged += OnSegmentsChanged;
        _outlookCalendar.EventsUpdated += OnOutlookEventsUpdated;

        QuickAddCommand = new RelayCommand(QuickAdd);
        SaveCommand = new RelayCommand(SaveTask, () => SelectedTask != null);
        ReopenCommand = new RelayCommand(ReopenSelectedTask, () => SelectedTask?.Status == TaskStatus.Done);
        DoneCommand = new RelayCommand(MarkSelectedTaskDone, () => SelectedTask != null);
        StartTimerCommand = new RelayCommand(StartTimer, () => SelectedTask != null);
        StopTimerCommand = new RelayCommand(StopTimer, () => SelectedTask != null);
        Add15Command = new RelayCommand(() => AdjustBookedMinutes(15), () => SelectedTask != null);
        Add30Command = new RelayCommand(() => AdjustBookedMinutes(30), () => SelectedTask != null);
        Add60Command = new RelayCommand(() => AdjustBookedMinutes(60), () => SelectedTask != null);
        Subtract15Command = new RelayCommand(() => AdjustBookedMinutes(-15), () => SelectedTask != null);
        Subtract30Command = new RelayCommand(() => AdjustBookedMinutes(-30), () => SelectedTask != null);
        Subtract60Command = new RelayCommand(() => AdjustBookedMinutes(-60), () => SelectedTask != null);
        BookTimeInTicketSystemCommand = new RelayCommand(async () => await BookTimeInTicketSystemAsync(), () => HasZnunyTicket && !IsTicketBooking && !_hasUnresolvedTicketTimeBooking && UnbookedTicketSeconds > 0);
        RefreshTicketFieldOptionsCommand = new RelayCommand(async () => await RefreshTicketFieldOptionsAsync(), () => HasZnunyTicket && !IsTicketBooking);
        CheckTicketTimeBookingCommand = new RelayCommand<TicketTimeBooking>(async booking => await CheckTicketTimeBookingAsync(booking), booking => booking?.CanCheckStatus == true && !IsTicketBooking);
        RetryTicketTimeBookingCommand = new RelayCommand<TicketTimeBooking>(async booking => await RetryTicketTimeBookingAsync(booking), booking => booking?.CanRetry == true && !IsTicketBooking);
        ComeCommand = new RelayCommand(() => { _workDays.SetCome(DateTime.Now); Load(); });
        GoCommand = new RelayCommand(() => { _workDays.SetGo(DateTime.Now); Load(); });
        BreakStartCommand = new RelayCommand(() => { _workDays.StartBreak(DateTime.Today.ToString("yyyy-MM-dd")); Load(); });
        BreakEndCommand = new RelayCommand(() => { _workDays.EndBreak(DateTime.Today.ToString("yyyy-MM-dd")); Load(); });
        ManualSaveCommand = new RelayCommand(SaveManualDay);
        AddBreakRowCommand = new RelayCommand(() => BreakRows.Add(new BreakEditRow()));
        SaveMarkersCommand = new RelayCommand(SaveMarkers);
        SetDayTypeNormalCommand = new RelayCommand(() => SetDayType("Normal"));
        SetDayTypeAmCommand = new RelayCommand(() => SetDayType("AM"));
        SetDayTypeUlCommand = new RelayCommand(() => SetDayType("UL"));
        AddSegmentCommand = new RelayCommand(AddSegment, () => CanSaveNewSegment);
        ShowTodayTasksCommand = new RelayCommand(() => SelectedTaskScope = TodayTaskScope.Today);
        ShowCurrentTasksCommand = new RelayCommand(() => SelectedTaskScope = TodayTaskScope.Current);
        ShowCompletedTasksCommand = new RelayCommand(() => SelectedTaskScope = TodayTaskScope.Completed);

        SelectTaskCommand = new RelayCommand<TaskItem>(task => SelectedTask = task, task => task != null);
        StartTaskCommand = new RelayCommand<TaskItem>(StartTaskFromCard);
        StopTaskCommand = new RelayCommand<TaskItem>(task => OnCardTaskAction(task, _tasks.StopTimer));
        DoneTaskCommand = new RelayCommand<TaskItem>(task => OnCardTaskAction(task, _tasks.MarkDone));
        TogglePinTaskCommand = new RelayCommand<TaskItem>(TogglePinTask, task => task != null);
        OpenTicketUrlCommand = new RelayCommand<string>(OpenTicketUrl, url => !string.IsNullOrWhiteSpace(url));
        OpenAgendaOutlookEventCommand = new RelayCommand<OutlookCalendarEvent>(OpenAgendaOutlookEvent, outlookEvent => outlookEvent != null);
        OpenAgendaTeamsCommand = new RelayCommand<string>(OpenAgendaTeams, url => !string.IsNullOrWhiteSpace(url));
        SaveSegmentCommand = new RelayCommand<TaskSegment>(SaveSegment, seg => seg != null && seg.IsValid);
        DeleteSegmentCommand = new RelayCommand<TaskSegment>(DeleteSegment, seg => seg != null);
        DeleteSegmentOutlookCommand = new RelayCommand<TaskSegment>(DeleteSegmentOutlook, seg => seg != null && !string.IsNullOrWhiteSpace(seg.OutlookEntryId));

        _clock = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
        _clock.Tick += (_, _) => OnClockTick();
        _clock.Start();

        SelectedTaskScope = TodayTaskScope.Today;
        Load();
    }

    private void RaiseSegmentEditorState()
    {
        Raise(nameof(NewSegmentValidationHint));
        Raise(nameof(CanSaveNewSegment));
        EvaluateNewSegmentConflict();
        AddSegmentCommand.RaiseCanExecuteChanged();
    }

    private void RaiseCommandStates()
    {
        SaveCommand.RaiseCanExecuteChanged();
        ReopenCommand.RaiseCanExecuteChanged();
        DoneCommand.RaiseCanExecuteChanged();
        StartTimerCommand.RaiseCanExecuteChanged();
        StopTimerCommand.RaiseCanExecuteChanged();
        Add15Command.RaiseCanExecuteChanged();
        Add30Command.RaiseCanExecuteChanged();
        Add60Command.RaiseCanExecuteChanged();
        Subtract15Command.RaiseCanExecuteChanged();
        Subtract30Command.RaiseCanExecuteChanged();
        Subtract60Command.RaiseCanExecuteChanged();
        BookTimeInTicketSystemCommand.RaiseCanExecuteChanged();
        RefreshTicketFieldOptionsCommand.RaiseCanExecuteChanged();
        AddSegmentCommand.RaiseCanExecuteChanged();
        SaveSegmentCommand.RaiseCanExecuteChanged();
        DeleteSegmentCommand.RaiseCanExecuteChanged();
        DeleteSegmentOutlookCommand.RaiseCanExecuteChanged();
    }

    public void Refresh()
    {
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher != null && !dispatcher.CheckAccess())
        {
            dispatcher.BeginInvoke(new Action(Load));
            return;
        }

        Load();
    }

    private void OnSegmentsChanged()
    {
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher != null && !dispatcher.CheckAccess())
        {
            dispatcher.BeginInvoke(new Action(ApplyTaskFilters));
            return;
        }

        ApplyTaskFilters();
    }

    private void OnOutlookEventsUpdated()
    {
        var dispatcher = Application.Current?.Dispatcher;
        if (dispatcher != null && !dispatcher.CheckAccess())
        {
            dispatcher.BeginInvoke(new Action(RefreshTodayAgenda));
            return;
        }

        RefreshTodayAgenda();
    }

    private void Load()
    {
        var selectedId = SelectedTask?.Id;
        ApplyTaskFilters();

        SelectedTask = selectedId.HasValue
            ? TodayTasks.Concat(CurrentTasks).Concat(CompletedTasks).FirstOrDefault(t => t.Id == selectedId.Value)
            : DisplayedTasks.FirstOrDefault() ?? CurrentTasks.FirstOrDefault();

        var day = DateTime.Today.ToString("yyyy-MM-dd");
        var wd = _workDays.GetOrCreateDay(day);
        var breaks = _workDays.GetBreaks(day);

        BreakRows.Clear();
        foreach (var br in breaks)
            BreakRows.Add(new BreakEditRow { StartTime = br.StartLocal.ToString("HH:mm"), EndTime = br.EndLocal?.ToString("HH:mm") ?? string.Empty, Note = br.Note });
        if (BreakRows.Count == 0) BreakRows.Add(new BreakEditRow());

        DayType = wd.DayType;
        IsBr = wd.IsBr;
        IsHo = wd.IsHo;

        ComeTimeText = wd.ComeLocal?.ToString("HH:mm") ?? string.Empty;
        GoTimeText = wd.GoLocal?.ToString("HH:mm") ?? string.Empty;

        WorkDaySummary = $"Kommen: {Fmt(wd.ComeLocal)}   Gehen: {Fmt(wd.GoLocal)}   Typ: {wd.DayType}";
        TicketMinutesToday = _tasks.GetTicketMinutesForDay(DateTime.Today);
        TicketMinutesCurrentMonth = _tasks.GetMonthTicketMinutes(DateTime.Today);
        StatusMessage = _tasks.LastError;
        RaiseSegmentEditorState();
        RaiseCommandStates();
        UpdateTimerDisplay();
    }

    private void ApplyTaskFilters()
    {
        var all = _tasks.GetAllTasks();
        var localToday = _germanTime.GetLocalNow(_settings.Current.CalendarTimeZoneId).Date;
        var todayTaskIds = _tasks.GetTaskIdsWithSegmentsForRange(localToday, localToday.AddDays(1));

        var active = all.Where(t => t.Status != TaskStatus.Done).ToList();
        if (!string.IsNullOrWhiteSpace(TaskSearchText))
        {
            var q = TaskSearchText.Trim();
            active = active.Where(t => (t.Title?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false)
                                     || (t.Description?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false)
                                     || (t.TicketUrl?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false)).ToList();
        }

        var done = all.Where(t => t.Status == TaskStatus.Done).ToList();
        if (!string.IsNullOrWhiteSpace(CompletedTaskSearchText))
        {
            var q = CompletedTaskSearchText.Trim();
            done = done.Where(t => (t.Title?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false)
                                 || (t.Description?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false)
                                 || (t.TicketUrl?.Contains(q, StringComparison.OrdinalIgnoreCase) ?? false)).ToList();
        }

        TodayTasks.Clear();
        foreach (var task in active.Where(task => todayTaskIds.Contains(task.Id)))
            TodayTasks.Add(task);

        CurrentTasks.Clear();
        foreach (var task in active)
            CurrentTasks.Add(task);

        _currentTasksWithoutToday.Clear();
        foreach (var task in active.Where(task => !todayTaskIds.Contains(task.Id)))
            _currentTasksWithoutToday.Add(task);

        CompletedTasks.Clear();
        foreach (var t in done) CompletedTasks.Add(t);

        RefreshDisplayedTasks();
        RefreshTodayAgenda();
    }

    private void RefreshTodayAgenda()
    {
        var localToday = _germanTime.GetLocalNow(_settings.Current.CalendarTimeZoneId).Date;
        _agendaDate = localToday;
        var tomorrow = localToday.AddDays(1);
        var segments = _tasks.GetSegmentsForRange(localToday, tomorrow);
        var visibleTodayTasks = TodayTasks.ToDictionary(task => task.Id);
        var mirroredOutlookEntryIds = segments
            .Select(pair => pair.Segment.OutlookEntryId)
            .Where(entryId => !string.IsNullOrWhiteSpace(entryId))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var agenda = new List<TodayAgendaItem>();
        foreach (var (task, segment) in segments)
        {
            if (task.Status == TaskStatus.Done || !visibleTodayTasks.TryGetValue(task.Id, out var visibleTask))
                continue;

            agenda.Add(new TodayAgendaItem
            {
                Start = segment.StartLocal,
                End = segment.EndLocal,
                Title = visibleTask.Title,
                Task = visibleTask,
                Segment = segment
            });
        }

        foreach (var outlookEvent in _outlookCalendar.GetEvents(localToday, tomorrow))
        {
            if (outlookEvent.IsCancelled || IsMirroredTaskSegment(outlookEvent, mirroredOutlookEntryIds))
                continue;

            if (outlookEvent.IsAllDay
                && _settings.Current.OutlookInterpretAllDayAsMarkers
                && OutlookAllDayMarkerMapper.TryMapAllDayMarker(outlookEvent, out _) != null)
                continue;

            agenda.Add(new TodayAgendaItem
            {
                Start = outlookEvent.StartLocal,
                End = outlookEvent.EndLocal,
                Title = outlookEvent.Subject,
                Location = outlookEvent.Location,
                OutlookEvent = outlookEvent,
                IsAllDay = outlookEvent.IsAllDay
            });
        }

        TodayAgendaItems.Clear();
        foreach (var item in agenda
                     .OrderBy(item => item.IsAllDay ? 0 : 1)
                     .ThenBy(item => item.Start)
                     .ThenBy(item => item.IsTask ? 0 : 1)
                     .ThenBy(item => item.Title, StringComparer.CurrentCultureIgnoreCase))
            TodayAgendaItems.Add(item);
    }

    private static bool IsMirroredTaskSegment(OutlookCalendarEvent outlookEvent, HashSet<string> mirroredOutlookEntryIds)
    {
        return mirroredOutlookEntryIds.Contains(outlookEvent.EntryId)
               || mirroredOutlookEntryIds.Contains(outlookEvent.Id);
    }

    private void RefreshDisplayedTasks()
    {
        DisplayedTasks.Clear();
        var source = SelectedTaskScope == TodayTaskScope.Today ? TodayTasks : _currentTasksWithoutToday;
        IEnumerable<TaskItem> ordered = source;
        if (SelectedTaskScope == TodayTaskScope.Current)
        {
            var pinnedFirst = source.OrderByDescending(task => task.IsPinned);
            var sortByCreated = string.Equals(_settings.Current.CurrentTasksSortField, "Created", StringComparison.OrdinalIgnoreCase);
            ordered = _settings.Current.CurrentTasksSortDescending
                ? sortByCreated
                    ? pinnedFirst.ThenByDescending(task => task.CreatedUtc)
                    : pinnedFirst.ThenByDescending(task => task.UpdatedUtc)
                : sortByCreated
                    ? pinnedFirst.ThenBy(task => task.CreatedUtc)
                    : pinnedFirst.ThenBy(task => task.UpdatedUtc);
        }

        foreach (var task in ordered)
            DisplayedTasks.Add(task);
    }

    private void TogglePinTask(TaskItem? task)
    {
        if (task == null || SelectedTaskScope != TodayTaskScope.Current) return;
        _tasks.SetPinned(task, !task.IsPinned);
        RefreshDisplayedTasks();
    }

    private void SetDayType(string type)
    {
        DayType = type;
        SaveMarkers();
    }

    private void SaveMarkers()
    {
        _workDays.SetDayMarkers(DateTime.Today.ToString("yyyy-MM-dd"), DayType, IsBr, IsHo);
        Load();
    }

    private void LoadSegments()
    {
        Segments.Clear();
        if (SelectedTask == null) return;

        var orderedSegments = _tasks.GetSegments(SelectedTask.Id).OrderBy(s => s.StartLocal).ToList();
        for (var i = 0; i < orderedSegments.Count; i++)
        {
            var seg = orderedSegments[i];
            seg.DisplayIndex = i + 1;
            seg.OutlookStatus = string.IsNullOrWhiteSpace(seg.OutlookEntryId) ? "fehlt" : "vorhanden";
            Segments.Add(seg);
        }

        EvaluateNewSegmentConflict();
    }

    private void EvaluateNewSegmentConflict()
    {
        if (!_settings.Current.OutlookCalendarEnabled || !_settings.Current.OutlookConflictWarningsEnabled || NewSegmentDate == null)
        {
            NewSegmentConflictWarning = string.Empty;
            return;
        }

        if (!TimeSpan.TryParse(NewSegmentStartTime, out var start) || !TimeSpan.TryParse(NewSegmentEndTime, out var end) || start >= end)
        {
            NewSegmentConflictWarning = string.Empty;
            return;
        }

        var startLocal = NewSegmentDate.Value.Date + start;
        var endLocal = NewSegmentDate.Value.Date + end;
        var conflicts = GetOutlookConflicts(startLocal, endLocal);
        NewSegmentConflictWarning = conflicts.Count == 0
            ? string.Empty
            : $"Konflikt mit Outlook: {string.Join(", ", conflicts.Select(c => c.Subject).Distinct().Take(2))}";
    }

    private List<OutlookCalendarEvent> GetOutlookConflicts(DateTime startLocal, DateTime endLocal)
    {
        var events = _outlookCalendar.GetEvents(startLocal.Date.AddDays(-1), startLocal.Date.AddDays(2));
        return events
            .Where(e => endLocal > e.StartLocal && startLocal < e.EndLocal)
            .Where(e => !IsDerivedDayMarkerEvent(e))
            .ToList();
    }

    private bool IsDerivedDayMarkerEvent(OutlookCalendarEvent evt)
    {
        if (!_settings.Current.OutlookInterpretAllDayAsMarkers)
            return false;

        var duration = evt.EndLocal - evt.StartLocal;
        if (!IsMarkerEligibleAllDayEvent(evt, duration))
            return false;

        return OutlookAllDayMarkerMapper.TryMapAllDayMarker(evt, out _) != null;
    }

    private static bool IsMarkerEligibleAllDayEvent(OutlookCalendarEvent evt, TimeSpan duration)
    {
        _ = duration;
        return evt.IsAllDay;
    }

    private bool ConfirmConflictIfRequired(DateTime startLocal, DateTime endLocal)
    {
        if (!_settings.Current.OutlookCalendarEnabled || !_settings.Current.OutlookConflictWarningsEnabled)
            return true;

        var conflicts = GetOutlookConflicts(startLocal, endLocal);
        if (conflicts.Count == 0)
            return true;

        var msg = "Dieses Segment überschneidet sich mit Outlook-Terminen:\n- " +
                  string.Join("\n- ", conflicts.Take(3).Select(c => $"{c.Subject} ({c.StartLocal:HH:mm}-{c.EndLocal:HH:mm})")) +
                  "\n\nTrotzdem speichern?";
        return MessageBox.Show(msg, "Outlook Konflikt", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes;
    }

    private void QuickAdd()
    {
        var task = _tasks.ParseQuickAdd(QuickAddText);
        _tasks.CreateTask(task);
        QuickAddText = string.Empty;
        SelectedTaskScope = TodayTaskScope.Current;
        Load();
        SelectedTask = CurrentTasks.FirstOrDefault(t => t.Id == task.Id)
                    ?? CompletedTasks.FirstOrDefault(t => t.Id == task.Id)
                    ?? task;
    }

    private void SaveTask()
    {
        if (SelectedTask == null) return;
        _tasks.UpdateTask(SelectedTask);
        StatusMessage = "Task gespeichert.";
        Load();
    }

    private static DateTime BuildSegmentDateTime(DateTime day, string timeText)
    {
        if (!TimeSpan.TryParse(timeText, out var time))
            throw new InvalidOperationException("Zeitformat ungültig.");

        return day.Date + time;
    }

    private void AddSegment()
    {
        if (SelectedTask == null || !CanSaveNewSegment || NewSegmentDate == null) return;

        var segment = new TaskSegment
        {
            TaskId = SelectedTask.Id,
            StartLocal = BuildSegmentDateTime(NewSegmentDate.Value, NewSegmentStartTime),
            EndLocal = BuildSegmentDateTime(NewSegmentDate.Value, NewSegmentEndTime),
            Note = NewSegmentNote,
            OutlookEntryId = string.Empty
        };
        segment.PlannedMinutes = (int)(segment.EndLocal - segment.StartLocal).TotalMinutes;
        if (!ConfirmConflictIfRequired(segment.StartLocal, segment.EndLocal))
            return;

        _tasks.AddSegment(segment);
        var outlookStatus = SyncSegmentOutlookAutomatically(segment);
        ServiceLocator.Notifications.RefreshSchedule();
        StatusMessage = $"Segment hinzugefügt.{outlookStatus}";
        LoadSegments();
        RaiseCommandStates();
    }

    private void SaveSegment(TaskSegment? segment)
    {
        if (segment == null) return;
        if (!segment.IsValid)
        {
            StatusMessage = segment.ValidationHint;
            return;
        }

        segment.PlannedMinutes = (int)(segment.EndLocal - segment.StartLocal).TotalMinutes;
        if (!ConfirmConflictIfRequired(segment.StartLocal, segment.EndLocal))
            return;

        _tasks.UpdateSegment(segment);
        var outlookStatus = SyncSegmentOutlookAutomatically(segment);
        ServiceLocator.Notifications.RefreshSchedule();
        segment.OutlookStatus = string.IsNullOrWhiteSpace(segment.OutlookEntryId) ? "fehlt" : "vorhanden";
        StatusMessage = $"Segment gespeichert.{outlookStatus}";
        RaiseCommandStates();
    }

    private void DeleteSegment(TaskSegment? segment)
    {
        if (segment == null) return;

        if (!_tasks.DeleteSegmentOutlook(segment))
        {
            segment.OutlookStatus = "fehler";
        }

        _tasks.DeleteSegment(segment.Id);
        ServiceLocator.Notifications.RefreshSchedule();
        StatusMessage = "Segment gelöscht.";
        LoadSegments();
    }

    private string SyncSegmentOutlookAutomatically(TaskSegment segment)
    {
        if (SelectedTask == null)
            return string.Empty;

        if (!_settings.Current.OutlookSyncEnabled)
        {
            segment.OutlookStatus = string.IsNullOrWhiteSpace(segment.OutlookEntryId) ? "fehlt" : "vorhanden";
            return " Outlook Sync ist deaktiviert.";
        }

        if (!_tasks.SyncSegmentOutlook(segment, SelectedTask.Title, SelectedTask.Description, SelectedTask.TicketUrl))
        {
            segment.OutlookStatus = "fehler";
            return $" Outlook Sync Fehler: {_tasks.LastError}";
        }

        segment.OutlookStatus = "vorhanden";
        return " Outlook wurde synchronisiert.";
    }

    private void DeleteSegmentOutlook(TaskSegment? segment)
    {
        if (segment == null) return;

        if (!_tasks.DeleteSegmentOutlook(segment))
        {
            segment.OutlookStatus = "fehler";
            StatusMessage = _tasks.LastError;
            return;
        }

        segment.OutlookStatus = "fehlt";
        StatusMessage = "Outlook Blocker für Segment gelöscht.";
        RaiseCommandStates();
    }

    private void StartTaskFromCard(TaskItem? task)
    {
        if (task == null) return;
        SelectedTask = task;
        StartTimer();
    }

    private void StartTimer()
    {
        if (SelectedTask == null) return;

        var runningOther = CurrentTasks.FirstOrDefault(t => t.Status == TaskStatus.Running && t.Id != SelectedTask.Id);
        if (runningOther != null)
            _tasks.StopTimer(runningOther);

        _tasks.StartTimer(SelectedTask);
        Load();
    }

    private void StopTimer()
    {
        if (SelectedTask == null) return;
        _tasks.StopTimer(SelectedTask);
        Load();
    }

    private void AdjustBookedMinutes(int deltaMinutes)
    {
        if (SelectedTask == null) return;
        _tasks.AddTicketMinutes(SelectedTask, deltaMinutes);
        Load();
    }

    private void LoadTicketBookingHistory()
    {
        TicketTimeBookings.Clear();
        _successfullyTransferredSeconds = 0;
        _successfullyBookedMinutes = 0;
        _ticketTimeBookingBaselineSeconds = 0;
        _hasUnresolvedTicketTimeBooking = false;
        if (SelectedTask != null)
        {
            foreach (var booking in _tasks.GetAllTicketTimeBookings(SelectedTask.Id))
                TicketTimeBookings.Add(booking);
            var successful = TicketTimeBookings.Where(booking => booking.Status == "Succeeded").ToList();
            _hasUnresolvedTicketTimeBooking = TicketTimeBookings.Any(booking => booking.Status != "Succeeded");
            _successfullyTransferredSeconds = successful.Sum(booking => booking.SourceSeconds);
            _successfullyBookedMinutes = successful.Sum(booking => booking.BookedMinutes);
            _ticketTimeBookingBaselineSeconds = _tasks.GetTicketTimeBookingBaselineSeconds(SelectedTask.Id);
        }

        Raise(nameof(SuccessfullyTransferredSeconds));
        Raise(nameof(UnbookedTicketSeconds));
        Raise(nameof(UnbookedTicketTimeText));
        Raise(nameof(TransferredTicketTimeText));
        BookTimeInTicketSystemCommand.RaiseCanExecuteChanged();
        CheckTicketTimeBookingCommand.RaiseCanExecuteChanged();
        RetryTicketTimeBookingCommand.RaiseCanExecuteChanged();
    }

    private async Task LoadTicketBookingContextAsync(TaskItem? task)
    {
        CostCenterOptions.Clear();
        OrderOptions.Clear();
        SelectedCostCenter = null;
        SelectedOrder = null;
        TicketBookingInformation = string.Empty;
        if (task == null || !task.Tags.Contains("ZnunyTicketID:", StringComparison.OrdinalIgnoreCase))
            return;

        try
        {
            var context = await _ticketSystem.GetTicketBookingContextAsync(task);
            if (SelectedTask?.Id != task.Id) return;
            foreach (var option in context.CostCenterOptions) CostCenterOptions.Add(option);
            foreach (var option in context.OrderOptions) OrderOptions.Add(option);
            SelectedCostCenter = EnsureCurrentOption(CostCenterOptions, context.CostCenterValue);
            SelectedOrder = EnsureCurrentOption(OrderOptions, context.OrderValue);
            TicketBookingInformation = context.Information;
        }
        catch (Exception ex)
        {
            if (SelectedTask?.Id == task.Id)
                TicketBookingInformation = $"Ticketdaten konnten nicht geladen werden: {ex.Message}";
        }
    }

    private async Task RefreshTicketFieldOptionsAsync()
    {
        if (SelectedTask == null || IsTicketBooking) return;
        IsTicketBooking = true;
        try
        {
            _ticketSystem.InvalidateDynamicFieldOptionsCache();
            await LoadTicketBookingContextAsync(SelectedTask);
            StatusMessage = "Kostenstellen und Aufträge wurden neu geladen.";
        }
        finally
        {
            IsTicketBooking = false;
        }
    }

    private static TicketFieldOption? EnsureCurrentOption(ObservableCollection<TicketFieldOption> options, string currentValue)
    {
        if (string.IsNullOrWhiteSpace(currentValue))
            return options.FirstOrDefault(option => option.Key == "00000");
        var existing = options.FirstOrDefault(option => string.Equals(option.Key, currentValue, StringComparison.OrdinalIgnoreCase));
        if (existing != null) return existing;
        var current = new TicketFieldOption(currentValue, currentValue);
        options.Insert(0, current);
        return current;
    }

    private async Task BookTimeInTicketSystemAsync()
    {
        if (SelectedTask == null || IsTicketBooking) return;
        var task = SelectedTask;
        IsTicketBooking = true;
        try
        {
            if (task.Status == TaskStatus.Running)
                _tasks.StopTimer(task);
            UpdateTimerDisplay();
            var seconds = UnbookedTicketSeconds;
            var description = FirstDescriptionLine(task.Description);
            var result = await _ticketSystem.BookTimeAsync(
                task,
                seconds,
                description,
                SelectedCostCenter?.Key ?? string.Empty,
                SelectedOrder?.Key ?? string.Empty);
            StatusMessage = result.Message;
            LoadTicketBookingHistory();
            UpdateTimerDisplay();
        }
        finally
        {
            IsTicketBooking = false;
        }
    }

    private async Task CheckTicketTimeBookingAsync(TicketTimeBooking? booking)
    {
        if (SelectedTask == null || booking == null || IsTicketBooking) return;
        var task = SelectedTask;
        IsTicketBooking = true;
        try
        {
            var result = await _ticketSystem.CheckTicketTimeBookingAsync(task, booking);
            StatusMessage = result.Message;
            if (SelectedTask?.Id == task.Id)
            {
                LoadTicketBookingHistory();
                UpdateTimerDisplay();
            }
        }
        finally
        {
            IsTicketBooking = false;
        }
    }

    private async Task RetryTicketTimeBookingAsync(TicketTimeBooking? booking)
    {
        if (SelectedTask == null || booking == null || IsTicketBooking) return;
        var task = SelectedTask;
        IsTicketBooking = true;
        try
        {
            var result = await _ticketSystem.RetryTicketTimeBookingAsync(task, booking);
            StatusMessage = result.Message;
            if (SelectedTask?.Id == task.Id)
            {
                LoadTicketBookingHistory();
                UpdateTimerDisplay();
            }
        }
        finally
        {
            IsTicketBooking = false;
        }
    }

    private static string FirstDescriptionLine(string description)
    {
        var line = (description ?? string.Empty)
            .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .FirstOrDefault() ?? "Zeitbuchung";
        return line.Length <= 500 ? line : line[..500];
    }

    private void ReopenSelectedTask() { if (SelectedTask == null) return; _tasks.MarkPlanned(SelectedTask); Load(); }
    private void MarkSelectedTaskDone() { if (SelectedTask == null) return; _tasks.MarkDone(SelectedTask); Load(); }

    private void OpenTicketUrl(string? url)
    {
        if (string.IsNullOrWhiteSpace(url)) return;
        ServiceLocator.MainViewModel.NavigateToTicketSystem(url);
    }

    private void OpenAgendaOutlookEvent(OutlookCalendarEvent? outlookEvent)
    {
        if (outlookEvent == null || string.IsNullOrWhiteSpace(outlookEvent.Id))
            return;

        var opened = ServiceLocator.Outlook.OpenCalendarEvent(outlookEvent.Id);
        if (!opened.ok)
            StatusMessage = $"Outlook-Termin konnte nicht geöffnet werden: {opened.error}";
    }

    private void OpenAgendaTeams(string? url)
    {
        if (!UrlLauncher.TryOpen(url, out var error))
            StatusMessage = $"Teams-Link konnte nicht geöffnet werden: {error}";
    }

    private void SaveManualDay()
    {
        try
        {
            var day = DateTime.Today;
            var come = ParseLocalTime(day, ComeTimeText);
            var go = ParseLocalTime(day, GoTimeText);
            var breaks = BreakRows
                .Select(row => new { row, start = ParseLocalTime(day, row.StartTime) })
                .Where(x => x.start.HasValue)
                .Select(x => new BreakRecord { Day = day.ToString("yyyy-MM-dd"), StartLocal = x.start!.Value, EndLocal = ParseLocalTime(day, x.row.EndTime), Note = x.row.Note })
                .ToList();

            _workDays.SaveManualDay(day.ToString("yyyy-MM-dd"), come, go, breaks);
            _workDays.SetDayMarkers(day.ToString("yyyy-MM-dd"), DayType, IsBr, IsHo);
            Load();
        }
        catch (Exception ex) { StatusMessage = $"Manuelles Speichern fehlgeschlagen: {ex.Message}"; }
    }

    private static DateTime? ParseLocalTime(DateTime day, string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        if (DateTime.TryParse(text, out var full)) return full;
        if (TimeSpan.TryParse(text, out var time)) return day.Date + time;
        return null;
    }

    private static string Fmt(DateTime? dt) => dt?.ToString("HH:mm") ?? "--:--";

    private void OnCardTaskAction(TaskItem? task, Action<TaskItem> action) { if (task == null) return; SelectedTask = task; action(task); Load(); }
    private void WithTask(Action<TaskItem> action) { if (SelectedTask == null) { MessageBox.Show("Bitte zuerst eine Aufgabe auswählen."); return; } action(SelectedTask); Load(); }

    private void UpdateTimerDisplay()
    {
        var unbooked = TimeSpan.FromSeconds(UnbookedTicketSeconds);
        TimerDisplay = $"{(int)unbooked.TotalHours:00}:{unbooked.Minutes:00}:{unbooked.Seconds:00}";
        Raise(nameof(UnbookedTicketSeconds));
        Raise(nameof(UnbookedTicketTimeText));
        Raise(nameof(TransferredTicketTimeText));
        BookTimeInTicketSystemCommand.RaiseCanExecuteChanged();
    }

    private void OnClockTick()
    {
        UpdateTimerDisplay();
        var localToday = _germanTime.GetLocalNow(_settings.Current.CalendarTimeZoneId).Date;
        if (localToday == _agendaDate)
            return;

        ApplyTaskFilters();
        _ = _outlookCalendar.TriggerSyncAsync(localToday, localToday.AddDays(1), "today-agenda-day-change");
    }


    public bool NavigateToTask(Guid taskId)
    {
        Load();

        var inToday = TodayTasks.FirstOrDefault(t => t.Id == taskId);
        if (inToday != null)
        {
            SelectedTaskScope = TodayTaskScope.Today;
            SelectedTask = inToday;
            TaskBringIntoViewRequested?.Invoke(taskId);
            return true;
        }

        var inCurrent = CurrentTasks.FirstOrDefault(t => t.Id == taskId);
        if (inCurrent != null)
        {
            SelectedTaskScope = TodayTaskScope.Current;
            SelectedTask = inCurrent;
            TaskBringIntoViewRequested?.Invoke(taskId);
            return true;
        }

        var inCompleted = CompletedTasks.FirstOrDefault(t => t.Id == taskId);
        if (inCompleted != null)
        {
            SelectedTaskScope = TodayTaskScope.Completed;
            SelectedTask = inCompleted;
            TaskBringIntoViewRequested?.Invoke(taskId);
            return true;
        }

        var task = _tasks.GetAllTasks().FirstOrDefault(t => t.Id == taskId);
        if (task == null)
            return false;

        SelectedTaskScope = task.Status == TaskStatus.Done ? TodayTaskScope.Completed : TodayTaskScope.Current;
        Load();
        SelectedTask = TodayTasks.FirstOrDefault(t => t.Id == taskId)
                    ?? CurrentTasks.FirstOrDefault(t => t.Id == taskId)
                    ?? CompletedTasks.FirstOrDefault(t => t.Id == taskId)
                    ?? task;
        TaskBringIntoViewRequested?.Invoke(taskId);
        return SelectedTask != null && SelectedTask.Id == taskId;
    }

    public override string ToString() => Title;
}

public enum TodayTaskScope
{
    Today,
    Current,
    Completed
}

public class BreakEditRow : ObservableObject
{
    private string _startTime = string.Empty;
    public string StartTime { get => _startTime; set => Set(ref _startTime, value); }
    private string _endTime = string.Empty;
    public string EndTime { get => _endTime; set => Set(ref _endTime, value); }
    private string _note = string.Empty;
    public string Note { get => _note; set => Set(ref _note, value); }
}
