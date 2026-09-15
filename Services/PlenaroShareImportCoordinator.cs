namespace TaskTool.Services;

/// Coordinates the bounded Outlook cache range needed for receiving shared tasks.
public sealed class PlenaroShareImportCoordinator : IDisposable
{
    internal const int PastDays = 30;
    internal const int FutureDays = 180;
    private readonly OutlookCalendarService _calendar;
    private readonly PlenaroShareImportService _importer;
    private readonly SettingsService _settings;
    private readonly LoggerService _logger;
    private readonly Func<DateTime> _today;
    private bool _disposed;

    public PlenaroShareImportCoordinator(OutlookCalendarService calendar, PlenaroShareImportService importer,
        SettingsService settings, LoggerService logger, Func<DateTime>? today = null)
    {
        _calendar = calendar;
        _importer = importer;
        _settings = settings;
        _logger = logger;
        _today = today ?? (() => DateTime.Today);
        _calendar.EventsUpdated += OnEventsUpdated;
        _settings.SettingsChanged += OnSettingsChanged;
    }

    public Task StartAsync() => EnsureAndImportAsync("share-startup");

    private void OnSettingsChanged() => _ = EnsureAndImportAsync("share-settings-changed");

    private void OnEventsUpdated()
    {
        if (_disposed || !_settings.Current.OutlookCalendarEnabled)
            return;
        var (from, to) = GetRange();
        // Import only complete ranges. A smaller sync may publish EventsUpdated while
        // the share range is already queued; importing is deferred until that queue finishes.
        if (_calendar.IsRangeAvailable(from, to))
            _importer.Import(_calendar.GetEvents(from, to));
    }

    private async Task EnsureAndImportAsync(string reason)
    {
        if (_disposed)
            return;
        if (!_settings.Current.OutlookCalendarEnabled)
        {
            _logger.Warning("Outlook-Kalenderintegration muss zum Empfangen geteilter Aufgaben aktiviert sein.");
            return;
        }
        var (from, to) = GetRange();
        await _calendar.EnsureRangeAvailableAsync(from, to, reason);
        if (!_disposed && _calendar.IsRangeAvailable(from, to))
            _importer.Import(_calendar.GetEvents(from, to));
    }

    internal (DateTime FromInclusive, DateTime ToExclusive) GetRange()
    {
        var today = _today().Date;
        return (today.AddDays(-PastDays), today.AddDays(FutureDays + 1));
    }

    public void Dispose()
    {
        _disposed = true;
        _calendar.EventsUpdated -= OnEventsUpdated;
        _settings.SettingsChanged -= OnSettingsChanged;
    }
}
