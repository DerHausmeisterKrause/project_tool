using System.Threading.Tasks;
using System.Linq;
using System.Collections.Generic;
using System;
using System.Diagnostics;
using System.Windows.Threading;
using TaskTool.Models;

namespace TaskTool.Services;

public class OutlookCalendarService : IDisposable
{
    private readonly LoggerService _logger;
    private readonly SettingsService _settings;
    private readonly OutlookInteropService _outlook;
    private readonly DispatcherTimer _timer;
    private readonly EventHandler _timerHandler;
    private readonly object _syncLock = new();

    private List<OutlookCalendarEvent> _cache = new();
    private bool _isSyncing;
    private bool _disposed;

    public event Action? EventsUpdated;

    public string LastError { get; private set; } = string.Empty;
    public DateTime? LastSyncAtLocal { get; private set; }

    public OutlookCalendarService(LoggerService logger, SettingsService settings, OutlookInteropService outlook)
    {
        _logger = logger;
        _settings = settings;
        _outlook = outlook;
        _timer = new DispatcherTimer();
        _timerHandler = async (_, _) => await TriggerSyncAsync("timer");
        _timer.Tick += _timerHandler;
        HandleSettingsChanged();
    }

    public IReadOnlyList<OutlookCalendarEvent> GetEvents(DateTime fromLocal, DateTime toLocal)
    {
        lock (_syncLock)
        {
            return _cache
                .Where(e => e.EndLocal > fromLocal && e.StartLocal < toLocal)
                .OrderBy(e => e.StartLocal)
                .ToList();
        }
    }

    public async Task TriggerSyncAsync(string reason = "manual")
    {
        if (_disposed || !_settings.Current.OutlookCalendarEnabled)
            return;

        if (_isSyncing)
            return;

        _isSyncing = true;
        var sw = Stopwatch.StartNew();

        try
        {
            var from = DateTime.Today.AddDays(-Math.Max(0, _settings.Current.OutlookCalendarRangePastDays));
            var to = DateTime.Today.AddDays(Math.Max(1, _settings.Current.OutlookCalendarRangeFutureDays + 1));

            _logger.Info($"[OutlookCalendarSync] Start reason={reason} from={from:O} to={to:O}");
            var result = await Task.Run(() => _outlook.GetCalendarEvents(from, to));
            if (!result.ok)
            {
                LastError = result.error;
                _logger.Warn($"[OutlookCalendarSync] Failed: {result.error}");
                return;
            }

            lock (_syncLock)
                _cache = result.events.OrderBy(e => e.StartLocal).ToList();

            LastError = string.Empty;
            LastSyncAtLocal = DateTime.Now;
            _logger.Info($"[OutlookCalendarSync] End events={_cache.Count} durationMs={sw.ElapsedMilliseconds}");
            EventsUpdated?.Invoke();
        }
        catch (Exception ex)
        {
            LastError = ex.Message;
            _logger.Error($"[OutlookCalendarSync] Exception: {ex}");
        }
        finally
        {
            _isSyncing = false;
        }
    }

    public void HandleSettingsChanged()
    {
        if (_disposed)
            return;

        var periodic = _settings.Current.OutlookCalendarEnabled
                       && string.Equals(_settings.Current.OutlookCalendarSyncMode, "Periodic", StringComparison.OrdinalIgnoreCase);

        _timer.Stop();
        if (periodic)
        {
            var intervalMin = Math.Max(1, _settings.Current.OutlookCalendarSyncIntervalMinutes);
            _timer.Interval = TimeSpan.FromMinutes(intervalMin);
            _timer.Start();
        }

        if (_settings.Current.OutlookCalendarEnabled)
            _ = TriggerSyncAsync("settings-changed");
        else
        {
            lock (_syncLock)
                _cache.Clear();
            EventsUpdated?.Invoke();
        }
    }

    public void Dispose()
    {
        _disposed = true;
        _timer.Stop();
        _timer.Tick -= _timerHandler;
    }
}
