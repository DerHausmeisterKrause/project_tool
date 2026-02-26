using System.Windows.Threading;
using TaskTool.Models;
using TaskTool.Views;

namespace TaskTool.Services;

public class NotificationService
{
    private readonly LoggerService _logger;
    private readonly SettingsService _settings;
    private readonly TaskService _tasks;
    private readonly DispatcherTimer _timer;
    private readonly Dictionary<Guid, DateTime> _snoozedUntil = new();
    private readonly HashSet<Guid> _shownForCurrentStart = new();

    public NotificationService(LoggerService logger, SettingsService settings, TaskService tasks)
    {
        _logger = logger;
        _settings = settings;
        _tasks = tasks;
        _timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(30) };
        _timer.Tick += (_, _) => CheckReminders();
        _timer.Start();
    }

    private void CheckReminders()
    {
        try
        {
            var now = DateTime.Now;
            var lead = _settings.Current.ReminderLeadMinutes;
            var upcoming = _tasks.GetUpcomingTasks(now, now.AddMinutes(lead + 1));
            foreach (var task in upcoming.Where(t => t.StartLocal.HasValue))
            {
                if (_snoozedUntil.TryGetValue(task.Id, out var snooze) && now < snooze) continue;
                var keyReached = _shownForCurrentStart.Contains(task.Id);
                if (keyReached) continue;

                var win = new ReminderWindow(task.Title, task.StartLocal!.Value);
                win.SnoozeRequested += (_, _) => _snoozedUntil[task.Id] = now.AddMinutes(5);
                win.Show();
                _shownForCurrentStart.Add(task.Id);
            }
        }
        catch (Exception ex)
        {
            _logger.Error($"Reminder check failed: {ex.Message}");
        }
    }
}
