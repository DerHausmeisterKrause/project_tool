using System.Windows;
using System.Windows.Threading;
using TaskTool.Infrastructure;
using TaskTool.Models;
using TaskTool.Services;
using TaskStatus = TaskTool.Models.TaskStatus;

namespace TaskTool.ViewModels;

public class DynamicIslandViewModel : ObservableObject
{
    private readonly DispatcherTimer _tick;
    private DateTime _nextFocusRefreshAt = DateTime.MinValue;

    private bool _isExpanded;
    public bool IsExpanded { get => _isExpanded; set => Set(ref _isExpanded, value); }

    private string _taskTitle = "Kein Fokus";
    public string TaskTitle { get => _taskTitle; set => Set(ref _taskTitle, value); }

    private string _timerDisplay = "00:00:00";
    public string TimerDisplay { get => _timerDisplay; set => Set(ref _timerDisplay, value); }

    private string _statusText = "Stopped";
    public string StatusText { get => _statusText; set => Set(ref _statusText, value); }

    private bool _hasTicketUrl;
    public bool HasTicketUrl { get => _hasTicketUrl; set => Set(ref _hasTicketUrl, value); }

    private bool _isRunning;
    public bool IsRunning { get => _isRunning; set => Set(ref _isRunning, value); }

    private string _nextFocusText = "Kein geplanter Fokus";
    public string NextFocusText { get => _nextFocusText; set => Set(ref _nextFocusText, value); }

    public bool ShowNextFocus => !IsRunning;

    public RelayCommand ToggleExpandCommand { get; }
    public RelayCommand StartStopCommand { get; }
    public RelayCommand OpenTaskCommand { get; }
    public RelayCommand OpenLinkCommand { get; }
    public RelayCommand NewTaskCommand { get; }

    public DynamicIslandViewModel()
    {
        ToggleExpandCommand = new RelayCommand(() => IsExpanded = !IsExpanded);
        StartStopCommand = new RelayCommand(StartStop);
        OpenTaskCommand = new RelayCommand(OpenTaskInApp);
        OpenLinkCommand = new RelayCommand(OpenTaskLink);
        NewTaskCommand = new RelayCommand(OpenNewTaskEntry);

        _tick = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1) };
        _tick.Tick += (_, _) => Refresh();
        _tick.Start();

        Refresh();
    }

    public void Stop() => _tick.Stop();

    private TodayViewModel Today => ServiceLocator.MainViewModel.TodayViewModel;

    private TaskItem? ResolveTask()
    {
        var running = Today.CurrentTasks.FirstOrDefault(t => t.Status == TaskStatus.Running)
                      ?? Today.CompletedTasks.FirstOrDefault(t => t.Status == TaskStatus.Running);
        return running ?? Today.SelectedTask ?? Today.CurrentTasks.FirstOrDefault() ?? Today.CompletedTasks.FirstOrDefault();
    }

    private void EnsureSelected(TaskItem task)
    {
        if (Today.SelectedTask?.Id == task.Id) return;
        Today.SelectedTask = Today.CurrentTasks.FirstOrDefault(t => t.Id == task.Id)
                            ?? Today.CompletedTasks.FirstOrDefault(t => t.Id == task.Id)
                            ?? task;
    }

    public void Refresh()
    {
        var task = ResolveTask();
        if (task == null)
        {
            TaskTitle = "Kein Fokus";
            TimerDisplay = "00:00:00";
            StatusText = "Stopped";
            HasTicketUrl = false;
            IsRunning = false;
            Raise(nameof(ShowNextFocus));
            RefreshNextFocusIfNeeded();
            return;
        }

        EnsureSelected(task);
        TaskTitle = task.Title;
        TimerDisplay = Today.TimerDisplay;
        var wasRunning = IsRunning;
        IsRunning = task.Status == TaskStatus.Running;
        StatusText = IsRunning ? "Running" : "Stopped";
        HasTicketUrl = !string.IsNullOrWhiteSpace(task.TicketUrl);
        if (wasRunning != IsRunning) Raise(nameof(ShowNextFocus));

        RefreshNextFocusIfNeeded();
    }

    private void RefreshNextFocusIfNeeded()
    {
        if (DateTime.Now < _nextFocusRefreshAt) return;
        _nextFocusRefreshAt = DateTime.Now.AddSeconds(20);

        var now = DateTime.Now;
        var next = ServiceLocator.Tasks.GetSegmentsForRange(now, now.AddDays(30))
            .OrderBy(s => s.Segment.StartLocal)
            .FirstOrDefault();

        if (next.Task == null)
        {
            NextFocusText = "Kein geplanter Fokus";
            return;
        }

        NextFocusText = $"Nächster Fokus: {next.Segment.StartLocal:HH:mm} · {next.Task.Title}";
    }

    private void StartStop()
    {
        var task = ResolveTask();
        if (task == null) return;

        EnsureSelected(task);
        if (task.Status == TaskStatus.Running)
            Today.StopTimerCommand.Execute(null);
        else
            Today.StartTimerCommand.Execute(null);

        Refresh();
    }

    private void OpenTaskInApp()
    {
        var task = ResolveTask();
        if (task == null) return;

        ActivateMainWindow();
        ServiceLocator.MainViewModel.NavigateToTodayAndOpenTask(task.Id);
        IsExpanded = false;
    }

    private void OpenNewTaskEntry()
    {
        ActivateMainWindow();
        ServiceLocator.MainViewModel.NavigateToTodayAndFocusQuickAdd();
        IsExpanded = false;
    }

    private void OpenTaskLink()
    {
        var task = ResolveTask();
        if (task == null || string.IsNullOrWhiteSpace(task.TicketUrl)) return;

        Today.OpenTicketUrlCommand.Execute(task.TicketUrl);
        IsExpanded = false;
    }

    private static void ActivateMainWindow()
    {
        var mainWindow = Application.Current.MainWindow;
        if (mainWindow == null) return;

        if (mainWindow.WindowState == WindowState.Minimized)
            mainWindow.WindowState = WindowState.Normal;

        mainWindow.Show();
        mainWindow.Topmost = true;
        mainWindow.Topmost = false;
        mainWindow.Activate();
        mainWindow.Focus();
    }
}
