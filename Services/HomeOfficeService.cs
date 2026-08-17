using TaskTool.Models;
using TaskTool.ViewModels;

namespace TaskTool.Services;

public sealed record HomeOfficeOperationResult(bool Success, bool IsHomeOffice, string Message);

public sealed class HomeOfficeService
{
    public event Action? Changed;
    private readonly WorkDayService _workDays;
    private readonly SettingsService _settings;
    private readonly OutlookInteropService _outlook;
    private readonly OutlookCalendarService _calendar;
    private readonly LoggerService _logger;

    public HomeOfficeService(WorkDayService workDays, SettingsService settings, OutlookInteropService outlook, OutlookCalendarService calendar, LoggerService logger)
    {
        _workDays = workDays;
        _settings = settings;
        _outlook = outlook;
        _calendar = calendar;
        _logger = logger;
    }

    public async Task<HomeOfficeOperationResult> SubmitAsync(DateTime day)
    {
        day = day.Date;
        var recipients = new[] { _settings.Current.HomeOfficeMailRecipient1, _settings.Current.HomeOfficeMailRecipient2 }
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => value.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();
        if (recipients.Length == 0)
            return new(false, false, "Es ist kein Homeoffice-E-Mail-Empfänger konfiguriert. Bitte zuerst in den Einstellungen mindestens einen Empfänger hinterlegen.");

        _logger.Info($"[HomeOffice] action=submit date={day:yyyy-MM-dd} confirmed=true");
        try
        {
            return await Task.Run(async () =>
            {
            var workDay = _workDays.GetOrCreateDay(day.ToString("yyyy-MM-dd"));
            var entryId = workDay.HomeOfficeOutlookEntryId;
            var createdNew = false;
            if (string.IsNullOrWhiteSpace(entryId))
            {
                var search = FindHomeOfficeCandidates(day);
                if (!search.ok)
                    return new HomeOfficeOperationResult(false, false, $"Outlook-Kalender konnte nicht auf vorhandene Homeoffice-Termine geprüft werden: {search.error}");
                if (search.events.Count > 1)
                    return new HomeOfficeOperationResult(false, false, "Mehrere Homeoffice-Termine wurden gefunden. Es wurde kein weiterer Termin erstellt.");
                entryId = search.events.SingleOrDefault()?.EntryId ?? string.Empty;
                createdNew = string.IsNullOrWhiteSpace(entryId);
            }

            var appointment = _outlook.UpsertHomeOfficeAppointment(entryId, day);
            _logger.Info($"[HomeOfficeOutlook] action={(createdNew ? "create" : "reuse")} date={day:yyyy-MM-dd} success={appointment.ok.ToString().ToLowerInvariant()} entryId='{appointment.entryId}'");
            if (!appointment.ok)
                return new HomeOfficeOperationResult(false, false, $"Outlook-Homeoffice-Termin konnte nicht erstellt werden: {appointment.error}");

            try
            {
                _workDays.SetHomeOfficeState(day.ToString("yyyy-MM-dd"), true, appointment.entryId);
                Changed?.Invoke();
            }
            catch (Exception ex)
            {
                if (createdNew) _outlook.DeleteBlock(appointment.entryId, ignoreSyncDisabled: true);
                _logger.Error($"[HomeOffice] action=submit date={day:yyyy-MM-dd} localSaveFailed='{ex.Message}'");
                return new HomeOfficeOperationResult(false, false, "Homeoffice konnte lokal nicht gespeichert werden.");
            }

            var mail = _outlook.SendHomeOfficeMail(day, recipients);
            _logger.Info($"[HomeOfficeMail] date={day:yyyy-MM-dd} recipientCount={recipients.Length} sendResult={(mail.ok ? "success" : "failed")}");
            await _calendar.TriggerSyncAsync(day, day.AddDays(1), "homeoffice-submit");
            return mail.ok
                ? new HomeOfficeOperationResult(true, true, "Homeoffice wurde eingereicht.")
                : new HomeOfficeOperationResult(false, true, $"Homeoffice wurde gespeichert, aber die E-Mail konnte nicht eindeutig versendet werden: {mail.error}");
            });
        }
        catch (Exception ex)
        {
            _logger.Error($"[HomeOffice] action=submit date={day:yyyy-MM-dd} failed='{ex.Message}'");
            return new(false, false, $"Homeoffice konnte nicht eingereicht werden: {ex.Message}");
        }
    }

    public async Task<HomeOfficeOperationResult> RemoveAsync(DateTime day)
    {
        day = day.Date;
        try
        {
            return await Task.Run(async () =>
            {
            var workDay = _workDays.GetOrCreateDay(day.ToString("yyyy-MM-dd"));
            var entryId = workDay.HomeOfficeOutlookEntryId;
            if (string.IsNullOrWhiteSpace(entryId))
            {
                var search = FindHomeOfficeCandidates(day);
                if (!search.ok)
                    return new HomeOfficeOperationResult(false, true, $"Outlook-Kalender konnte nicht auf den Homeoffice-Termin geprüft werden: {search.error}");
                if (search.events.Count > 1)
                    return new HomeOfficeOperationResult(false, true, "Mehrere Homeoffice-Termine wurden gefunden. Aus Sicherheitsgründen wurde keiner gelöscht.");
                entryId = search.events.SingleOrDefault()?.EntryId ?? string.Empty;
            }

            var removed = string.IsNullOrWhiteSpace(entryId) ? (ok: true, error: string.Empty) : _outlook.DeleteBlock(entryId, ignoreSyncDisabled: true);
            if (!removed.ok)
                return new HomeOfficeOperationResult(false, true, $"Outlook-Homeoffice-Termin konnte nicht entfernt werden: {removed.error}");

            try
            {
                _workDays.SetHomeOfficeState(day.ToString("yyyy-MM-dd"), false, string.Empty);
                Changed?.Invoke();
            }
            catch (Exception ex) { return new HomeOfficeOperationResult(false, true, $"Lokaler Homeoffice-Marker konnte nicht entfernt werden: {ex.Message}"); }
            _logger.Info($"[HomeOffice] action=remove date={day:yyyy-MM-dd} outlookRemoved=true");
            await _calendar.TriggerSyncAsync(day, day.AddDays(1), "homeoffice-remove");
            return new HomeOfficeOperationResult(true, false, "Homeoffice wurde entfernt.");
            });
        }
        catch (Exception ex)
        {
            _logger.Error($"[HomeOffice] action=remove date={day:yyyy-MM-dd} failed='{ex.Message}'");
            return new(false, true, $"Homeoffice konnte nicht entfernt werden: {ex.Message}");
        }
    }

    private (bool ok, List<OutlookCalendarEvent> events, string error) FindHomeOfficeCandidates(DateTime day)
    {
        var fetched = _outlook.GetCalendarEvents(day.Date, day.Date.AddDays(1), ignoreCalendarDisabled: true);
        if (!fetched.ok) return (false, new List<OutlookCalendarEvent>(), fetched.error);
        var events = fetched.events
            .Where(item => item.IsAllDay && item.StartLocal.Date == day.Date)
            .Where(item => OutlookAllDayMarkerMapper.TryMapAllDayMarker(item, out _) == "HO")
            .ToList();
        return (true, events, string.Empty);
    }
}
