using System;
using System.Linq;
using System.Windows;
using System.Windows.Threading;
using TaskTool.Models;
using TaskStatus = TaskTool.Models.TaskStatus;
using TaskTool.Views;

namespace TaskTool.Services;

public enum ReminderKind
{
    Lead,
    Start
}

public class NotificationService
{
    private readonly LoggerService _logger;
    private readonly SettingsService _settings;
    private readonly TaskService _tasks;
    private readonly DispatcherTimer _timer;
    private readonly HashSet<string> _firedKeys = new();

    private DateTime _lastCheck;
    private DynamicIslandWindow? _islandWindow;

    public NotificationService(LoggerService logger, SettingsService settings, TaskService tasks)
    {
        _logger = logger;
        _settings = settings;
        _tasks = tasks;
        _lastCheck = DateTime.Now;
        _timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(20) };
        _timer.Tick += (_, _) => CheckReminders();
        _timer.Start();
        EnsureIslandWindow();

        _logger.Info("Notification scheduler started.");
    }


    private void EnsureIslandWindow()
    {
        Application.Current.Dispatcher.BeginInvoke(() =>
        {
            if (_islandWindow != null) return;
            _islandWindow = new DynamicIslandWindow();
            _islandWindow.Closed += (_, _) => _islandWindow = null;
            _islandWindow.Show();
        }, DispatcherPriority.Background);
    }

    public void RefreshSchedule()
    {
        _lastCheck = DateTime.Now;
        TrimFiredKeys();
        _logger.Info("Notification scheduler refreshed.");
    }

    public void ShowTestNotification()
    {
        try
        {
            var task = _tasks.GetAllTasks().FirstOrDefault(t => t.Status != TaskStatus.Done)
                       ?? _tasks.GetAllTasks().FirstOrDefault();

            if (task == null)
            {
                ShowOverlay(Guid.Empty, "Test Benachrichtigung (keine Aufgabe vorhanden)", ReminderKind.Lead);
                return;
            }

            var lead = Math.Max(1, _settings.Current.ReminderLeadMinutes);
            ShowOverlay(task.Id, $"{task.Title} in {lead} Minuten", ReminderKind.Lead);
        }
        catch (Exception ex)
        {
            _logger.Error($"Test notification failed: {ex}");
        }
    }

    private void CheckReminders()
    {
        try
        {
            var now = DateTime.Now;
            var lead = _settings.Current.ReminderLeadMinutes;

            var from = _lastCheck.AddHours(-1);
            var to = now.AddDays(2);
            var segments = _tasks.GetSegmentsForRange(from, to)
                .Where(x => x.Segment.StartLocal > now.AddHours(-1))
                .ToList();

            foreach (var (task, segment) in segments)
            {
                var start = segment.StartLocal;

                if (lead > 0)
                {
                    var leadTime = start.AddMinutes(-lead);
                    if (ShouldFire(leadTime, now))
                    {
                        var minutesLeft = Math.Max(1, (int)Math.Ceiling((start - now).TotalMinutes));
                        var text = $"{task.Title} in {minutesLeft} Minuten";
                        Fire(task.Id, segment.Id, start, ReminderKind.Lead, text);
                    }
                }

                if (ShouldFire(start, now))
                {
                    Fire(task.Id, segment.Id, start, ReminderKind.Start, $"{task.Title} jetzt");
                }
            }

            _lastCheck = now;
            TrimFiredKeys();
        }
        catch (Exception ex)
        {
            _logger.Error($"Reminder check failed: {ex}");
        }
    }

    private bool ShouldFire(DateTime triggerTime, DateTime now)
    {
        if (triggerTime < DateTime.Now.AddHours(-12)) return false;
        if (triggerTime > now) return false;
        return triggerTime > _lastCheck;
    }

    private void Fire(Guid taskId, long segmentId, DateTime start, ReminderKind kind, string text)
    {
        var key = BuildKey(segmentId, start, kind);
        if (!_firedKeys.Add(key)) return;

        _logger.Info($"Reminder fired: {key}");
        ShowOverlay(taskId, text, kind);
    }

    private void ShowOverlay(Guid taskId, string text, ReminderKind kind)
    {
        Application.Current.Dispatcher.BeginInvoke(() =>
        {
            var overlay = new ReminderWindow(text, kind, taskId);
            overlay.NotificationClicked += (_, id) => ActivateMainWindowAndOpenTask(id);
            overlay.Show();
        }, DispatcherPriority.Background);
    }

    private static string BuildKey(long segmentId, DateTime start, ReminderKind kind)
        => $"{segmentId}:{start:O}:{kind}";

    private void TrimFiredKeys()
    {
        if (_firedKeys.Count <= 2000) return;
        _firedKeys.Clear();
    }

    private static void ActivateMainWindowAndOpenTask(Guid taskId)
    {
        var mainWindow = Application.Current.MainWindow;
        if (mainWindow == null)
            return;

        if (mainWindow.WindowState == WindowState.Minimized)
            mainWindow.WindowState = WindowState.Normal;

        mainWindow.Show();
        mainWindow.Topmost = true;
        mainWindow.Topmost = false;
        mainWindow.Activate();
        mainWindow.Focus();

        if (taskId != Guid.Empty)
            ServiceLocator.MainViewModel.NavigateToTodayAndOpenTask(taskId);
    }
}
