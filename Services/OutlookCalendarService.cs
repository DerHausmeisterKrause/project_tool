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
            var visible = new List<OutlookCalendarEvent>();
            foreach (var e in _cache)
            {
                var overlap = e.StartLocal < toLocal && e.EndLocal > fromLocal;
                _logger.Info($"[OutlookRangeCheck] subject='{e.Subject}' start={e.StartLocal:O} end={e.EndLocal:O} isAllDay={e.IsAllDay} fromInclusive={fromLocal:O} toExclusive={toLocal:O} overlap={overlap}");
                if (overlap)
                {
                    visible.Add(e);
                }
                else
                {
                    _logger.Info($"[OutlookEventFiltered] subject='{e.Subject}' reason=FilteredByTimeRange start={e.StartLocal:O} end={e.EndLocal:O} fromInclusive={fromLocal:O} toExclusive={toLocal:O} overlap={overlap}");
                }
            }

            return visible.OrderBy(e => e.StartLocal).ToList();
        }
    }

    public Task TriggerSyncAsync(string reason = "manual")
    {
        var from = DateTime.Today.AddDays(-Math.Max(0, _settings.Current.OutlookCalendarRangePastDays));
        var to = DateTime.Today.AddDays(Math.Max(1, _settings.Current.OutlookCalendarRangeFutureDays + 1));
        return TriggerSyncAsync(from, to, reason);
    }

    public async Task TriggerSyncAsync(DateTime fromInclusiveLocal, DateTime toExclusiveLocal, string reason)
    {
        if (_disposed || !_settings.Current.OutlookCalendarEnabled)
            return;

        if (_isSyncing)
            return;

        _isSyncing = true;
        var sw = Stopwatch.StartNew();

        try
        {
            var from = fromInclusiveLocal;
            var to = toExclusiveLocal;

            _logger.Info($"[OutlookCalendarSync] Start reason={reason} from={from:O} to={to:O}");
            var result = await Task.Run(() => _outlook.GetCalendarEvents(from, to));
            if (!result.ok)
            {
                LastError = result.error;
                _logger.Error($"[OutlookCalendarSync] Failed: {result.error}");
                return;
            }

            foreach (var e in result.events)
            {
                _logger.Info($"[OutlookFetchedEvent] subject='{e.Subject}' start={e.StartLocal:O} end={e.EndLocal:O} isAllDay={e.IsAllDay} entryId='{e.EntryId}'");
                _logger.Info($"[OutlookRawEvent] subject='{e.Subject}' start={e.StartLocal:O} end={e.EndLocal:O} isAllDay={e.IsAllDay} busyStatus='{e.BusyStatus}' sensitivity='{e.Sensitivity}' isPrivate={e.IsPrivate} isRecurring={e.IsRecurring} isInstance={e.IsInstance} meetingStatus='{e.MeetingStatus}' messageClass='{e.MessageClass}' isCancelled={e.IsCancelled} categories='{e.Categories}' location='{e.Location}' calendar='{e.CalendarName}' entryId='{e.EntryId}' iCalUId='{e.ICalUId}'");
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
