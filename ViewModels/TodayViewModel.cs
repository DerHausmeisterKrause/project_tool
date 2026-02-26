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
    private readonly DispatcherTimer _clock;

    public string Title => "Heute";
    public ObservableCollection<TaskItem> CurrentTasks { get; } = new();
    public ObservableCollection<TaskItem> CompletedTasks { get; } = new();
    public ObservableCollection<BreakEditRow> BreakRows { get; } = new();
    public ObservableCollection<TaskSegment> Segments { get; } = new();
    public ObservableCollection<string> TimeOptions { get; } = new(Enumerable.Range(0, 96).Select(i => TimeSpan.FromMinutes(i * 15).ToString(@"hh\:mm")));

    private bool _freezeTimerDisplay;
    private string _frozenTimerDisplay = "00:00:00";

    private TaskItem? _selectedTask;
    public TaskItem? SelectedTask
    {
        get => _selectedTask;
        set
        {
            if (Set(ref _selectedTask, value))
            {
                _freezeTimerDisplay = false;
                LoadSegments();
                Raise(nameof(IsTaskSelected));
                RaiseCommandStates();
                UpdateTimerDisplay();
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

    private bool _showCompletedTasks;
    public bool ShowCompletedTasks
    {
        get => _showCompletedTasks;
        set
        {
            if (Set(ref _showCompletedTasks, value))
            {
                Raise(nameof(ShowCurrentTasks));
                Raise(nameof(ShowCompletedTaskList));
            }
        }
    }

    public bool ShowCurrentTasks => !ShowCompletedTasks;
    public bool ShowCompletedTaskList => ShowCompletedTasks;

    private string _statusMessage = string.Empty;
    public string StatusMessage { get => _statusMessage; set => Set(ref _statusMessage, value); }

    private string _workDaySummary = string.Empty;
    public string WorkDaySummary { get => _workDaySummary; set => Set(ref _workDaySummary, value); }

    private string _todayTotals = string.Empty;
    public string TodayTotals { get => _todayTotals; set => Set(ref _todayTotals, value); }

    private string _comeTimeText = string.Empty;
    public string ComeTimeText { get => _comeTimeText; set => Set(ref _comeTimeText, value); }

    private string _goTimeText = string.Empty;
    public string GoTimeText { get => _goTimeText; set => Set(ref _goTimeText, value); }

    private string _timerDisplay = "00:00:00";
    public string TimerDisplay { get => _timerDisplay; set => Set(ref _timerDisplay, value); }

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
    public RelayCommand PauseTimerCommand { get; }
    public RelayCommand StopTimerCommand { get; }
    public RelayCommand Add15Command { get; }
    public RelayCommand Add30Command { get; }
    public RelayCommand Add60Command { get; }
    public RelayCommand Subtract15Command { get; }
    public RelayCommand Subtract30Command { get; }
    public RelayCommand Subtract60Command { get; }
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
    public RelayCommand SyncAllSegmentsCommand { get; }
    public RelayCommand ShowCurrentTasksCommand { get; }
    public RelayCommand ShowCompletedTasksCommand { get; }
    public RelayCommand TestOutlookConnectionCommand { get; }

    public RelayCommand<TaskItem> SelectTaskCommand { get; }
    public RelayCommand<TaskItem> StartTaskCommand { get; }
    public RelayCommand<TaskItem> PauseTaskCommand { get; }
    public RelayCommand<TaskItem> StopTaskCommand { get; }
    public RelayCommand<TaskItem> DoneTaskCommand { get; }
    public RelayCommand<string> OpenTicketUrlCommand { get; }
    public RelayCommand<TaskSegment> SaveSegmentCommand { get; }
    public RelayCommand<TaskSegment> DeleteSegmentCommand { get; }
    public RelayCommand<TaskSegment> SyncSegmentOutlookCommand { get; }
    public RelayCommand<TaskSegment> DeleteSegmentOutlookCommand { get; }

    private string _dayType = "Normal";
    public string DayType { get => _dayType; set => Set(ref _dayType, value); }

    private bool _isBr;
    public bool IsBr { get => _isBr; set => Set(ref _isBr, value); }

    private bool _isHo;
    public bool IsHo { get => _isHo; set => Set(ref _isHo, value); }

    public TodayViewModel(TaskService tasks, WorkDayService workDays, SettingsService settings)
    {
        _tasks = tasks;
        _workDays = workDays;
        _settings = settings;

        QuickAddCommand = new RelayCommand(QuickAdd);
        SaveCommand = new RelayCommand(SaveTask, () => SelectedTask != null);
        ReopenCommand = new RelayCommand(ReopenTask, () => SelectedTask?.Status == TaskStatus.Done);
        DoneCommand = new RelayCommand(MarkDone, () => SelectedTask != null);
        StartTimerCommand = new RelayCommand(StartTimer, () => SelectedTask != null);
        PauseTimerCommand = new RelayCommand(PauseTimer, () => SelectedTask != null);
        StopTimerCommand = new RelayCommand(StopTimer, () => SelectedTask != null);
        Add15Command = new RelayCommand(() => AdjustBookedMinutes(15), () => SelectedTask != null);
        Add30Command = new RelayCommand(() => AdjustBookedMinutes(30), () => SelectedTask != null);
        Add60Command = new RelayCommand(() => AdjustBookedMinutes(60), () => SelectedTask != null);
        Subtract15Command = new RelayCommand(() => AdjustBookedMinutes(-15), () => SelectedTask != null);
        Subtract30Command = new RelayCommand(() => AdjustBookedMinutes(-30), () => SelectedTask != null);
        Subtract60Command = new RelayCommand(() => AdjustBookedMinutes(-60), () => SelectedTask != null);
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
        SyncAllSegmentsCommand = new RelayCommand(SyncAllSegments, () => SelectedTask != null);
        ShowCurrentTasksCommand = new RelayCommand(() => ShowCompletedTasks = false);
        ShowCompletedTasksCommand = new RelayCommand(() => ShowCompletedTasks = true);
        TestOutlookConnectionCommand = new RelayCommand(TestOutlookConnection);

        SelectTaskCommand = new RelayCommand<TaskItem>(task => SelectedTask = task, task => task != null);
        StartTaskCommand = new RelayCommand<TaskItem>(task => OnCardTaskAction(task, _tasks.StartTimer));
        PauseTaskCommand = new RelayCommand<TaskItem>(task => OnCardTaskAction(task, _tasks.PauseTimer));
        StopTaskCommand = new RelayCommand<TaskItem>(task => OnCardTaskAction(task, _tasks.StopTimer));
        DoneTaskCommand = new RelayCommand<TaskItem>(task => OnCardTaskAction(task, _tasks.MarkDone));
        OpenTicketUrlCommand = new RelayCommand<string>(OpenTicketUrl, url => !string.IsNullOrWhiteSpace(url));
        SaveSegmentCommand = new RelayCommand<TaskSegment>(SaveSegment, seg => seg != null && seg.IsValid);
        DeleteSegmentCommand = new RelayCommand<TaskSegment>(DeleteSegment, seg => seg != null);
        SyncSegmentOutlookCommand = new RelayCommand<TaskSegment>(SyncSegmentOutlook, seg => seg != null && seg.Id > 0);
        DeleteSegmentOutlookCommand = new RelayCommand<TaskSegment>(DeleteSegmentOutlook, seg => seg != null && !string.IsNullOrWhiteSpace(seg.OutlookEntryId));

        _clock = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
        _clock.Tick += (_, _) => UpdateTimerDisplay();
        _clock.Start();

        ShowCompletedTasks = false;
        Load();
    }

    private void RaiseSegmentEditorState()
    {
        Raise(nameof(NewSegmentValidationHint));
        Raise(nameof(CanSaveNewSegment));
        AddSegmentCommand.RaiseCanExecuteChanged();
    }

    private void RaiseCommandStates()
    {
        SaveCommand.RaiseCanExecuteChanged();
        ReopenCommand.RaiseCanExecuteChanged();
        DoneCommand.RaiseCanExecuteChanged();
        StartTimerCommand.RaiseCanExecuteChanged();
        PauseTimerCommand.RaiseCanExecuteChanged();
        StopTimerCommand.RaiseCanExecuteChanged();
        Add15Command.RaiseCanExecuteChanged();
        Add30Command.RaiseCanExecuteChanged();
        Add60Command.RaiseCanExecuteChanged();
        Subtract15Command.RaiseCanExecuteChanged();
        Subtract30Command.RaiseCanExecuteChanged();
        Subtract60Command.RaiseCanExecuteChanged();
        AddSegmentCommand.RaiseCanExecuteChanged();
        SyncAllSegmentsCommand.RaiseCanExecuteChanged();
        SaveSegmentCommand.RaiseCanExecuteChanged();
        DeleteSegmentCommand.RaiseCanExecuteChanged();
        SyncSegmentOutlookCommand.RaiseCanExecuteChanged();
        DeleteSegmentOutlookCommand.RaiseCanExecuteChanged();
    }

    private void Load()
    {
        var selectedId = SelectedTask?.Id;
        ApplyTaskFilters();

        SelectedTask = selectedId.HasValue
            ? CurrentTasks.Concat(CompletedTasks).FirstOrDefault(t => t.Id == selectedId.Value)
            : CurrentTasks.FirstOrDefault();

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

        var presence = (wd.ComeLocal.HasValue && wd.GoLocal.HasValue) ? wd.GoLocal.Value - wd.ComeLocal.Value : TimeSpan.Zero;
        var pause = breaks.Where(b => b.EndLocal.HasValue).Aggregate(TimeSpan.Zero, (acc, b) => acc + (b.EndLocal!.Value - b.StartLocal));
        var net = presence - pause;
        var target = (wd.DayType == "UL" || wd.DayType == "AM") ? 0 : _settings.Current.GetTargetMinutes(DateTime.Today.DayOfWeek);
        var overtime = (int)net.TotalMinutes - target;
        var monthOvertime = CalculateMonthOvertime();

        WorkDaySummary = $"Kommen: {Fmt(wd.ComeLocal)}   Gehen: {Fmt(wd.GoLocal)}   Typ: {wd.DayType}";
        TodayTotals = $"Soll: {FmtMin(target)} | Ist: {FmtMin((int)net.TotalMinutes)} | Ü heute: {FmtMin(overtime)} | Ü Monat: {FmtMin(monthOvertime)}";
        StatusMessage = _tasks.LastError;
        RaiseSegmentEditorState();
        RaiseCommandStates();
        UpdateTimerDisplay();
    }

    private void ApplyTaskFilters()
    {
        var all = _tasks.GetAllTasks();

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

        CurrentTasks.Clear();
        foreach (var t in active) CurrentTasks.Add(t);

        CompletedTasks.Clear();
        foreach (var t in done) CompletedTasks.Add(t);
    }

    private int CalculateMonthOvertime()
    {
        var first = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
        var last = first.AddMonths(1).AddDays(-1);
        var days = _workDays.GetWorkDaysInRange(first, last).ToDictionary(d => d.Day, d => d);
        var total = 0;
        for (var d = first; d <= last; d = d.AddDays(1))
        {
            var key = d.ToString("yyyy-MM-dd");
            var wd = days.ContainsKey(key) ? days[key] : new WorkDayRecord { Day = key, DayType = "Normal" };
            var target = (wd.DayType == "UL" || wd.DayType == "AM") ? 0 : _settings.Current.GetTargetMinutes(d.DayOfWeek);
            var netMin = 0;
            if (wd.ComeLocal.HasValue && wd.GoLocal.HasValue)
            {
                var breaks = _workDays.GetBreaks(key);
                var pause = breaks.Where(b => b.EndLocal.HasValue).Sum(b => (int)(b.EndLocal!.Value - b.StartLocal).TotalMinutes);
                netMin = (int)(wd.GoLocal.Value - wd.ComeLocal.Value).TotalMinutes - pause;
            }
            total += netMin - target;
        }
        return total;
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
    }

    private void QuickAdd()
    {
        var task = _tasks.ParseQuickAdd(QuickAddText);
        _tasks.CreateTask(task);
        QuickAddText = string.Empty;
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
        _tasks.AddSegment(segment);
        StatusMessage = "Segment hinzugefügt.";
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
        _tasks.UpdateSegment(segment);
        segment.OutlookStatus = string.IsNullOrWhiteSpace(segment.OutlookEntryId) ? "fehlt" : "vorhanden";
        StatusMessage = "Segment gespeichert.";
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
        StatusMessage = "Segment gelöscht.";
        LoadSegments();
    }

    private void SyncSegmentOutlook(TaskSegment? segment)
    {
        if (segment == null || SelectedTask == null) return;

        if (!_tasks.SyncSegmentOutlook(segment, SelectedTask.Title, SelectedTask.Description, SelectedTask.TicketUrl))
        {
            segment.OutlookStatus = "fehler";
            StatusMessage = _tasks.LastError;
            return;
        }

        segment.OutlookStatus = "vorhanden";
        StatusMessage = "Outlook Blocker für Segment synchronisiert.";
        RaiseCommandStates();
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

    private void SyncAllSegments()
    {
        if (SelectedTask == null) return;

        var errors = 0;
        foreach (var seg in Segments)
        {
            if (!_tasks.SyncSegmentOutlook(seg, SelectedTask.Title, SelectedTask.Description, SelectedTask.TicketUrl))
            {
                errors++;
                seg.OutlookStatus = "fehler";
                continue;
            }

            seg.OutlookStatus = "vorhanden";
        }

        StatusMessage = errors == 0 ? "Alle Segmente wurden mit Outlook synchronisiert." : $"{errors} Segment(e) mit Fehler synchronisiert.";
        RaiseCommandStates();
    }


    private void StartTimer()
    {
        if (SelectedTask == null) return;
        _freezeTimerDisplay = false;
        _tasks.StartTimer(SelectedTask);
        Load();
    }

    private void PauseTimer()
    {
        if (SelectedTask == null) return;
        var snapshot = TimerDisplay;
        _tasks.PauseTimer(SelectedTask);
        Load();
        _freezeTimerDisplay = true;
        _frozenTimerDisplay = snapshot;
        TimerDisplay = _frozenTimerDisplay;
    }

    private void StopTimer()
    {
        if (SelectedTask == null) return;
        var snapshot = TimerDisplay;
        _tasks.StopTimer(SelectedTask);
        Load();
        _freezeTimerDisplay = true;
        _frozenTimerDisplay = snapshot;
        TimerDisplay = _frozenTimerDisplay;
    }

    private void AdjustBookedMinutes(int deltaMinutes)
    {
        if (SelectedTask == null) return;
        _freezeTimerDisplay = false;
        _tasks.AddTicketMinutes(SelectedTask, deltaMinutes);
        Load();
    }

    private void ReopenTask() { if (SelectedTask == null) return; _tasks.MarkPlanned(SelectedTask); Load(); }
    private void MarkDone() { if (SelectedTask == null) return; _tasks.MarkDone(SelectedTask); Load(); }

    private void TestOutlookConnection()
    {
        var ok = _tasks.TestOutlookConnection();
        if (ok)
        {
            StatusMessage = "Outlook Verbindungstest erfolgreich.";
            MessageBox.Show(StatusMessage, "Outlook Test", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var hex = ExtractHResultHex(_tasks.LastError);
        StatusMessage = $"Outlook Test fehlgeschlagen. Details in logs.txt: {hex}";
        MessageBox.Show(StatusMessage, "Outlook Test", MessageBoxButton.OK, MessageBoxImage.Error);
    }


    private void OpenTicketUrl(string? url)
    {
        if (!UrlLauncher.TryOpen(url, out var error))
        {
            StatusMessage = error;
        }
    }

    private static string ExtractHResultHex(string error)
    {
        if (string.IsNullOrWhiteSpace(error)) return "0x00000000";
        var marker = "0x";
        var idx = error.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (idx < 0) return "0x00000000";
        var end = idx + 2;
        while (end < error.Length && Uri.IsHexDigit(error[end])) end++;
        return error[idx..end];
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
    private static string FmtMin(int mins) => $"{mins / 60}h {Math.Abs(mins % 60):00}m";

    private void OnCardTaskAction(TaskItem? task, Action<TaskItem> action) { if (task == null) return; SelectedTask = task; action(task); Load(); }
    private void WithTask(Action<TaskItem> action) { if (SelectedTask == null) { MessageBox.Show("Bitte zuerst eine Aufgabe auswählen."); return; } action(SelectedTask); Load(); }

    private void UpdateTimerDisplay()
    {
        if (SelectedTask == null) { TimerDisplay = "00:00:00"; return; }
        if (_freezeTimerDisplay) { TimerDisplay = _frozenTimerDisplay; return; }

        var booked = TimeSpan.FromMinutes(Math.Max(0, SelectedTask.TicketMinutesBooked));
        var runningPart = _tasks.GetOpenSessionDuration(SelectedTask.Id);
        var total = booked + runningPart;
        TimerDisplay = $"{(int)total.TotalHours:00}:{total.Minutes:00}:{total.Seconds:00}";
    }

    public override string ToString() => Title;
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
