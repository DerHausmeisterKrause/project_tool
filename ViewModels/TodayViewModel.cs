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

    public ObservableCollection<int> DurationOptions { get; } = new() { 15, 30, 45, 60, 90, 120, 180 };
    public ObservableCollection<int> Hours { get; } = new(Enumerable.Range(0, 24));
    public ObservableCollection<int> Minutes { get; } = new(new[] { 0, 5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55 });

    private TaskItem? _selectedTask;
    public TaskItem? SelectedTask
    {
        get => _selectedTask;
        set
        {
            if (Set(ref _selectedTask, value))
            {
                LoadTaskPlanning();
                LoadSegments();
                Raise(nameof(IsTaskSelected));
                Raise(nameof(OutlookBlockerButtonText));
                Raise(nameof(OutlookBlockerCanSync));
                Raise(nameof(IsSchedulingValid));
                Raise(nameof(OutlookSyncHint));
                Raise(nameof(ComputedEndText));
                RaiseCommandStates();
                UpdateTimerDisplay();
            }
        }
    }

    public bool IsTaskSelected => SelectedTask != null;
    public string DateTimeFormat => string.IsNullOrWhiteSpace(_settings.Current.DateTimeFormat) ? "yyyy-MM-dd HH:mm" : _settings.Current.DateTimeFormat;
    public bool AllowMultiDaySegments => _settings.Current.AllowMultiDayTaskPlanning;

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

    private DateTime? _selectedStartDate = DateTime.Today;
    public DateTime? SelectedStartDate { get => _selectedStartDate; set { if (Set(ref _selectedStartDate, value)) { Raise(nameof(ComputedEndText)); Raise(nameof(OutlookBlockerCanSync)); Raise(nameof(IsSchedulingValid)); Raise(nameof(OutlookSyncHint)); RaiseCommandStates(); } } }

    private int _selectedHour = DateTime.Now.Hour;
    public int SelectedHour { get => _selectedHour; set { if (Set(ref _selectedHour, value)) { Raise(nameof(ComputedEndText)); Raise(nameof(OutlookBlockerCanSync)); Raise(nameof(IsSchedulingValid)); Raise(nameof(OutlookSyncHint)); RaiseCommandStates(); } } }

    private int _selectedMinute;
    public int SelectedMinute { get => _selectedMinute; set { if (Set(ref _selectedMinute, value)) { Raise(nameof(ComputedEndText)); Raise(nameof(OutlookBlockerCanSync)); Raise(nameof(IsSchedulingValid)); Raise(nameof(OutlookSyncHint)); RaiseCommandStates(); } } }

    private int _selectedDurationMinutes = 30;
    public int SelectedDurationMinutes { get => _selectedDurationMinutes; set { if (Set(ref _selectedDurationMinutes, value)) { Raise(nameof(ComputedEndText)); Raise(nameof(OutlookBlockerCanSync)); Raise(nameof(IsSchedulingValid)); Raise(nameof(OutlookSyncHint)); RaiseCommandStates(); } } }

    private string _manualDurationMinutesText = string.Empty;
    public string ManualDurationMinutesText { get => _manualDurationMinutesText; set { if (Set(ref _manualDurationMinutesText, value)) { Raise(nameof(ComputedEndText)); Raise(nameof(OutlookBlockerCanSync)); Raise(nameof(IsSchedulingValid)); Raise(nameof(OutlookSyncHint)); RaiseCommandStates(); } } }

    private string _dayType = "Normal";
    public string DayType { get => _dayType; set => Set(ref _dayType, value); }

    private bool _isBr;
    public bool IsBr { get => _isBr; set => Set(ref _isBr, value); }

    private bool _isHo;
    public bool IsHo { get => _isHo; set => Set(ref _isHo, value); }

    public string ComputedEndText
    {
        get
        {
            if (!OutlookBlockerCanSync || !TryBuildStart(out var start)) return "-";
            return start.AddMinutes(GetDurationMinutes()).ToString(DateTimeFormat);
        }
    }

    public bool OutlookBlockerCanSync => SelectedTask != null && IsSchedulingValid;
    public bool IsSchedulingValid => TryBuildStart(out var start) && start > DateTime.MinValue && GetDurationMinutes() > 0;
    public string OutlookSyncHint => IsSchedulingValid ? "" : "Bitte gültigen Start + Dauer setzen (Ende > Start).";
    public string OutlookBlockerButtonText => string.IsNullOrWhiteSpace(SelectedTask?.OutlookEntryId) ? "Outlook Blocker erstellen" : "Outlook Blocker aktualisieren";

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
    public RelayCommand ComeCommand { get; }
    public RelayCommand GoCommand { get; }
    public RelayCommand BreakStartCommand { get; }
    public RelayCommand BreakEndCommand { get; }
    public RelayCommand ManualSaveCommand { get; }
    public RelayCommand AddBreakRowCommand { get; }
    public RelayCommand SyncOutlookBlockerCommand { get; }
    public RelayCommand DeleteOutlookBlockerCommand { get; }
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

    public TodayViewModel(TaskService tasks, WorkDayService workDays, SettingsService settings)
    {
        _tasks = tasks;
        _workDays = workDays;
        _settings = settings;

        QuickAddCommand = new RelayCommand(QuickAdd);
        SaveCommand = new RelayCommand(SaveTask, () => SelectedTask != null);
        ReopenCommand = new RelayCommand(ReopenTask, () => SelectedTask?.Status == TaskStatus.Done);
        DoneCommand = new RelayCommand(MarkDone, () => SelectedTask != null);
        StartTimerCommand = new RelayCommand(() => WithTask(_tasks.StartTimer), () => SelectedTask != null);
        PauseTimerCommand = new RelayCommand(() => WithTask(_tasks.PauseTimer), () => SelectedTask != null);
        StopTimerCommand = new RelayCommand(() => WithTask(_tasks.StopTimer), () => SelectedTask != null);
        Add15Command = new RelayCommand(() => WithTask(t => _tasks.AddTicketMinutes(t, 15)), () => SelectedTask != null);
        Add30Command = new RelayCommand(() => WithTask(t => _tasks.AddTicketMinutes(t, 30)), () => SelectedTask != null);
        Add60Command = new RelayCommand(() => WithTask(t => _tasks.AddTicketMinutes(t, 60)), () => SelectedTask != null);
        ComeCommand = new RelayCommand(() => { _workDays.SetCome(DateTime.Now); Load(); });
        GoCommand = new RelayCommand(() => { _workDays.SetGo(DateTime.Now); Load(); });
        BreakStartCommand = new RelayCommand(() => { _workDays.StartBreak(DateTime.Today.ToString("yyyy-MM-dd")); Load(); });
        BreakEndCommand = new RelayCommand(() => { _workDays.EndBreak(DateTime.Today.ToString("yyyy-MM-dd")); Load(); });
        ManualSaveCommand = new RelayCommand(SaveManualDay);
        AddBreakRowCommand = new RelayCommand(() => BreakRows.Add(new BreakEditRow()));
        SyncOutlookBlockerCommand = new RelayCommand(SyncOutlookBlocker, () => SelectedTask != null && OutlookBlockerCanSync);
        DeleteOutlookBlockerCommand = new RelayCommand(DeleteOutlookBlocker, () => SelectedTask != null);
        SaveMarkersCommand = new RelayCommand(SaveMarkers);
        SetDayTypeNormalCommand = new RelayCommand(() => SetDayType("Normal"));
        SetDayTypeAmCommand = new RelayCommand(() => SetDayType("AM"));
        SetDayTypeUlCommand = new RelayCommand(() => SetDayType("UL"));
        AddSegmentCommand = new RelayCommand(AddSegment, () => SelectedTask != null && AllowMultiDaySegments);
        SyncAllSegmentsCommand = new RelayCommand(SyncAllSegments, () => SelectedTask != null && AllowMultiDaySegments);
        ShowCurrentTasksCommand = new RelayCommand(() => ShowCompletedTasks = false);
        ShowCompletedTasksCommand = new RelayCommand(() => ShowCompletedTasks = true);
        TestOutlookConnectionCommand = new RelayCommand(TestOutlookConnection);

        SelectTaskCommand = new RelayCommand<TaskItem>(task => SelectedTask = task, task => task != null);
        StartTaskCommand = new RelayCommand<TaskItem>(task => OnCardTaskAction(task, _tasks.StartTimer));
        PauseTaskCommand = new RelayCommand<TaskItem>(task => OnCardTaskAction(task, _tasks.PauseTimer));
        StopTaskCommand = new RelayCommand<TaskItem>(task => OnCardTaskAction(task, _tasks.StopTimer));
        DoneTaskCommand = new RelayCommand<TaskItem>(task => OnCardTaskAction(task, _tasks.MarkDone));

        _clock = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
        _clock.Tick += (_, _) => UpdateTimerDisplay();
        _clock.Start();

        ShowCompletedTasks = false;
        Load();
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
        SyncOutlookBlockerCommand.RaiseCanExecuteChanged();
        DeleteOutlookBlockerCommand.RaiseCanExecuteChanged();
        AddSegmentCommand.RaiseCanExecuteChanged();
        SyncAllSegmentsCommand.RaiseCanExecuteChanged();
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

    private void LoadTaskPlanning()
    {
        if (SelectedTask?.StartLocal is not DateTime start) return;
        SelectedStartDate = start.Date;
        SelectedHour = start.Hour;
        SelectedMinute = (start.Minute / 5) * 5;
        if (SelectedTask.EndLocal.HasValue)
            SelectedDurationMinutes = Math.Max(15, (int)(SelectedTask.EndLocal.Value - start).TotalMinutes);
    }

    private void LoadSegments()
    {
        Segments.Clear();
        if (SelectedTask == null || !AllowMultiDaySegments) return;
        foreach (var seg in _tasks.GetSegments(SelectedTask.Id)) Segments.Add(seg);
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
        if (TryBuildStart(out var start))
        {
            SelectedTask.StartLocal = start;
            SelectedTask.EndLocal = start.AddMinutes(GetDurationMinutes());
        }
        _tasks.UpdateTask(SelectedTask);
        StatusMessage = "Task gespeichert.";
        Load();
    }

    private bool TryBuildStart(out DateTime start)
    {
        start = DateTime.MinValue;
        if (!SelectedStartDate.HasValue) return false;
        start = SelectedStartDate.Value.Date.AddHours(SelectedHour).AddMinutes(SelectedMinute);
        return true;
    }

    private int GetDurationMinutes() => int.TryParse(ManualDurationMinutesText, out var parsed) && parsed > 0 ? parsed : SelectedDurationMinutes;

    private void AddSegment()
    {
        if (SelectedTask == null || !TryBuildStart(out var start)) return;
        var dur = GetDurationMinutes();
        _tasks.AddSegment(new TaskSegment { TaskId = SelectedTask.Id, StartLocal = start, EndLocal = start.AddMinutes(dur), PlannedMinutes = dur });
        LoadSegments();
    }

    private void SyncAllSegments()
    {
        if (SelectedTask == null) return;
        foreach (var seg in Segments) _tasks.SyncSegmentOutlook(seg, SelectedTask.Title, SelectedTask.Description, SelectedTask.TicketUrl);
        StatusMessage = _tasks.LastError;
        LoadSegments();
    }

    private void ReopenTask() { if (SelectedTask == null) return; _tasks.MarkPlanned(SelectedTask); Load(); }
    private void MarkDone() { if (SelectedTask == null) return; _tasks.MarkDone(SelectedTask); Load(); }

    private void SyncOutlookBlocker()
    {
        if (SelectedTask == null || !TryBuildStart(out var start)) return;

        try
        {
            SelectedTask.StartLocal = start;
            SelectedTask.EndLocal = start.AddMinutes(GetDurationMinutes());
            var ok = _tasks.SyncOutlookBlocker(SelectedTask);
            if (!ok)
            {
                var hResultHex = ExtractHResultHex(_tasks.LastError);
                StatusMessage = $"Outlook Sync fehlgeschlagen. Details in logs.txt: {hResultHex}";
                MessageBox.Show(StatusMessage, "Outlook Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            StatusMessage = "Outlook Blocker synchronisiert.";
            Load();
        }
        catch (Exception ex)
        {
            var hResultHex = $"0x{ex.HResult:X8}";
            StatusMessage = $"Outlook Sync fehlgeschlagen. Details in logs.txt: {hResultHex}";
            MessageBox.Show(StatusMessage, "Outlook Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void DeleteOutlookBlocker()
    {
        if (SelectedTask == null) return;
        try
        {
            var ok = _tasks.DeleteOutlookBlocker(SelectedTask);
            if (!ok)
            {
                var hResultHex = ExtractHResultHex(_tasks.LastError);
                StatusMessage = $"Outlook Delete fehlgeschlagen. Details in logs.txt: {hResultHex}";
                MessageBox.Show(StatusMessage, "Outlook Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            StatusMessage = "Outlook Blocker gelöscht.";
            Load();
        }
        catch (Exception ex)
        {
            var hResultHex = $"0x{ex.HResult:X8}";
            StatusMessage = $"Outlook Delete fehlgeschlagen. Details in logs.txt: {hResultHex}";
            MessageBox.Show(StatusMessage, "Outlook Fehler", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }


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
        var elapsed = _tasks.GetTrackedDuration(SelectedTask.Id);
        TimerDisplay = $"{(int)elapsed.TotalHours:00}:{elapsed.Minutes:00}:{elapsed.Seconds:00}";
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
