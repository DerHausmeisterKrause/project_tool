using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using TaskTool.Models;

namespace TaskTool.Services;

public class OutlookInteropService
{
    private const int OlAppointmentItem = 1;
    private const int OlFolderCalendar = 9;
    private const int OlBusy = 2;

    private readonly LoggerService _logger;
    private readonly SettingsService _settings;

    public OutlookInteropService(LoggerService logger, SettingsService settings)
    {
        _logger = logger;
        _settings = settings;
    }

    public (bool ok, string entryId, string error) UpsertBlock(string? existingEntryId, string title, string body, DateTime start, DateTime end)
    {
        if (!_settings.Current.OutlookSyncEnabled)
            return (false, existingEntryId ?? string.Empty, "Outlook Sync ist deaktiviert.");

        if (string.IsNullOrWhiteSpace(title))
            return (false, existingEntryId ?? string.Empty, "Titel fehlt.");

        if (start == default || end == default || end <= start || start == DateTime.MinValue || end == DateTime.MinValue)
            return (false, existingEntryId ?? string.Empty, "Ungültiger Zeitraum: Ende muss nach Start liegen.");

        try
        {
            return ExecuteOnSta<(bool ok, string entryId, string error)>(() =>
            {
                var outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null)
                    return (false, existingEntryId ?? string.Empty, "Outlook nicht installiert (ProgID nicht gefunden).");

                object? app = null;
                object? ns = null;
                object? item = null;

                try
                {
                    app = CreateOrAttachOutlook(outlookType);
                    if (app == null)
                        return (false, existingEntryId ?? string.Empty, "Outlook konnte nicht gestartet/verbunden werden.");

                    dynamic appDyn = app;
                    ns = appDyn.GetNamespace("MAPI");
                    TryLogon(ns);

                    dynamic nsDyn = ns!;
                    _ = nsDyn.GetDefaultFolder(OlFolderCalendar);

                    if (!string.IsNullOrWhiteSpace(existingEntryId))
                    {
                        item = nsDyn.GetItemFromID(existingEntryId);
                    }
                    else
                    {
                        item = appDyn.CreateItem(OlAppointmentItem);
                    }

                    if (item == null)
                        return (false, existingEntryId ?? string.Empty, "Outlook Terminobjekt konnte nicht erstellt werden.");

                    dynamic itemDyn = item;
                    itemDyn.Subject = $"Fokus: {title}";
                    itemDyn.Body = body ?? string.Empty;
                    itemDyn.Start = start;
                    itemDyn.End = end;
                    itemDyn.BusyStatus = OlBusy;
                    itemDyn.ReminderSet = false;
                    itemDyn.Categories = string.IsNullOrWhiteSpace(_settings.Current.OutlookCategoryName)
                        ? "FocusBlock"
                        : _settings.Current.OutlookCategoryName;
                    itemDyn.Save();

                    var entryId = Convert.ToString(itemDyn.EntryID) ?? string.Empty;
                    return (true, entryId, string.Empty);
                }
                finally
                {
                    SafeReleaseComObject(item);
                    SafeReleaseComObject(ns);
                    SafeReleaseComObject(app);
                }
            });
        }
        catch (Exception ex)
        {
            _logger.Error(BuildOutlookExceptionLog("UpsertBlock", ex, start, end));
            return (false, existingEntryId ?? string.Empty, BuildUserFacingOutlookError(ex));
        }
    }

    public (bool ok, string error) DeleteBlock(string? entryId)
    {
        if (!_settings.Current.OutlookSyncEnabled || string.IsNullOrWhiteSpace(entryId))
            return (true, string.Empty);

        try
        {
            return ExecuteOnSta<(bool ok, string error)>(() =>
            {
                var outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null)
                    return (false, "Outlook nicht installiert (ProgID nicht gefunden).");

                object? app = null;
                object? ns = null;
                object? item = null;

                try
                {
                    app = CreateOrAttachOutlook(outlookType);
                    if (app == null)
                        return (false, "Outlook konnte nicht gestartet/verbunden werden.");

                    dynamic appDyn = app;
                    ns = appDyn.GetNamespace("MAPI");
                    TryLogon(ns);

                    dynamic nsDyn = ns!;
                    item = nsDyn.GetItemFromID(entryId);
                    if (item == null)
                        return (false, "Outlook Entry nicht gefunden.");

                    dynamic itemDyn = item;
                    itemDyn.Delete();
                    return (true, string.Empty);
                }
                finally
                {
                    SafeReleaseComObject(item);
                    SafeReleaseComObject(ns);
                    SafeReleaseComObject(app);
                }
            });
        }
        catch (Exception ex)
        {
            _logger.Error(BuildOutlookExceptionLog("DeleteBlock", ex, null, null));
            return (false, BuildUserFacingOutlookError(ex));
        }
    }


    public (bool ok, List<OutlookCalendarEvent> events, string error) GetCalendarEvents(DateTime fromLocal, DateTime toLocal)
    {
        if (!_settings.Current.OutlookCalendarEnabled)
            return (true, new List<OutlookCalendarEvent>(), string.Empty);

        if (toLocal <= fromLocal)
            return (false, new List<OutlookCalendarEvent>(), "Ungültiger Zeitraum für Kalenderabfrage.");

        try
        {
            return ExecuteOnSta(() =>
            {
                var outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null)
                    return (false, new List<OutlookCalendarEvent>(), "Outlook nicht installiert (ProgID nicht gefunden).");

                object? app = null;
                object? ns = null;
                object? folder = null;
                object? items = null;
                object? restricted = null;

                try
                {
                    app = CreateOrAttachOutlook(outlookType);
                    if (app == null)
                        return (false, new List<OutlookCalendarEvent>(), "Outlook konnte nicht gestartet/verbunden werden.");

                    dynamic appDyn = app;
                    ns = appDyn.GetNamespace("MAPI");
                    TryLogon(ns);
                    dynamic nsDyn = ns!;
                    folder = nsDyn.GetDefaultFolder(OlFolderCalendar);

                    dynamic folderDyn = folder!;
                    items = folderDyn.Items;
                    dynamic itemsDyn = items!;
                    itemsDyn.IncludeRecurrences = true;
                    itemsDyn.Sort("[Start]");

                    var fromFilter = fromLocal.ToString("MM/dd/yyyy HH:mm");
                    var toFilter = toLocal.ToString("MM/dd/yyyy HH:mm");
                    var filter = $"[Start] < '{toFilter}' AND [End] > '{fromFilter}'";
                    restricted = itemsDyn.Restrict(filter);

                    var events = new List<OutlookCalendarEvent>();
                    foreach (var raw in (System.Collections.IEnumerable)restricted!)
                    {
                        object? appointment = raw;
                        try
                        {
                            dynamic a = appointment!;
                            DateTime start = Convert.ToDateTime(a.Start).ToLocalTime();
                            DateTime end = Convert.ToDateTime(a.End).ToLocalTime();
                            if (end <= fromLocal || start >= toLocal)
                                continue;

                            var body = Convert.ToString(a.Body) ?? string.Empty;
                            var location = Convert.ToString(a.Location) ?? string.Empty;
                            var joinUrl = ExtractTeamsUrl(body, location);

                            events.Add(new OutlookCalendarEvent
                            {
                                Id = Convert.ToString(a.EntryID) ?? Guid.NewGuid().ToString("N"),
                                Subject = string.IsNullOrWhiteSpace(Convert.ToString(a.Subject)) ? "(Kein Betreff)" : Convert.ToString(a.Subject)!,
                                StartLocal = start,
                                EndLocal = end,
                                IsAllDay = Convert.ToBoolean(a.AllDayEvent),
                                Location = location,
                                Organizer = Convert.ToString(a.Organizer) ?? string.Empty,
                                BodyPreview = body.Length > 240 ? body[..240] : body,
                                OnlineMeetingJoinUrl = joinUrl,
                                Categories = Convert.ToString(a.Categories) ?? string.Empty
                            });
                        }
                        catch
                        {
                            // ignore non-appointment entries
                        }
                        finally
                        {
                            SafeReleaseComObject(appointment);
                        }
                    }

                    return (true, events, string.Empty);
                }
                finally
                {
                    SafeReleaseComObject(restricted);
                    SafeReleaseComObject(items);
                    SafeReleaseComObject(folder);
                    SafeReleaseComObject(ns);
                    SafeReleaseComObject(app);
                }
            });
        }
        catch (Exception ex)
        {
            _logger.Error(BuildOutlookExceptionLog("GetCalendarEvents", ex, fromLocal, toLocal));
            return (false, new List<OutlookCalendarEvent>(), BuildUserFacingOutlookError(ex));
        }
    }

    private static string ExtractTeamsUrl(string body, string location)
    {
        var pattern = @"https?://[^\s""']+";
        foreach (Match match in Regex.Matches($"{body}\n{location}", pattern, RegexOptions.IgnoreCase))
        {
            var url = match.Value.TrimEnd('.', ',', ';', ')');
            if (url.Contains("teams.microsoft.com", StringComparison.OrdinalIgnoreCase)
                || url.Contains("meetup-join", StringComparison.OrdinalIgnoreCase))
                return url;
        }

        return string.Empty;
    }

    public (bool ok, string error) TestConnection()
    {
        var start = DateTime.Now.AddMinutes(5);
        var end = start.AddMinutes(5);

        try
        {
            var upsert = UpsertBlock(string.Empty, "TaskTool Test", "Test appointment", start, end);
            if (!upsert.ok)
                return (false, upsert.error);

            var del = DeleteBlock(upsert.entryId);
            if (!del.ok)
                return (false, del.error);

            return (true, string.Empty);
        }
        catch (Exception ex)
        {
            _logger.Error(BuildOutlookExceptionLog("TestConnection", ex, start, end));
            return (false, BuildUserFacingOutlookError(ex));
        }
    }

    private static object? CreateOrAttachOutlook(Type outlookType)
    {
        return Activator.CreateInstance(outlookType);
    }

    private static void TryLogon(object nameSpace)
    {
        try
        {
            dynamic ns = nameSpace;
            ns.Logon("", "", false, false);
        }
        catch
        {
            // Often already logged on; safe to continue.
        }
    }

    private static void SafeReleaseComObject(object? comObject)
    {
        if (comObject == null)
            return;

        try
        {
            if (Marshal.IsComObject(comObject))
                Marshal.FinalReleaseComObject(comObject);
        }
        catch
        {
            // best effort cleanup only
        }
    }

    private static T ExecuteOnSta<T>(Func<T> action)
    {
        if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
            return action();

        T? result = default;
        Exception? exception = null;

        var thread = new Thread(() =>
        {
            try
            {
                result = action();
            }
            catch (Exception ex)
            {
                exception = ex;
            }
        });

        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        thread.Join();

        if (exception != null)
            throw new InvalidOperationException("Outlook COM Aufruf auf STA Thread fehlgeschlagen.", exception);

        return result!;
    }

    private static string BuildUserFacingOutlookError(Exception ex)
    {
        if (ex is FileNotFoundException || ex is TypeLoadException)
            return "Outlook-Interop konnte nicht geladen werden. Bitte Office/Outlook reparieren und App neu starten.";

        if (ex.Message.Contains("office, Version=", StringComparison.OrdinalIgnoreCase))
            return "Office Interop Assembly wurde nicht gefunden. Bitte Office/Outlook reparieren.";

        if (ex is COMException comEx)
        {
            if ((uint)comEx.HResult == 0x800401E3)
                return $"COM Fehler 0x{comEx.HResult:X8}: Kein aktives Outlook-Profil verfügbar.";

            if ((uint)comEx.HResult == 0x80070002)
                return $"COM Fehler 0x{comEx.HResult:X8}: Outlook-Dateien/Registrierung nicht gefunden.";

            return $"COM Fehler 0x{comEx.HResult:X8}: {comEx.Message}";
        }

        var message = string.IsNullOrWhiteSpace(ex.Message) ? "Unbekannter Outlook Fehler." : ex.Message;
        return $"{message} (0x{ex.HResult:X8})";
    }

    private static string BuildOutlookExceptionLog(string operation, Exception ex, DateTime? start, DateTime? end)
    {
        var sb = new StringBuilder();
        sb.AppendLine($"Outlook {operation} failed");
        sb.AppendLine($"ThreadId: {Environment.CurrentManagedThreadId}");
        sb.AppendLine($"ApartmentState: {Thread.CurrentThread.GetApartmentState()}");
        sb.AppendLine($"OutlookInstalled: {Type.GetTypeFromProgID("Outlook.Application") != null}");
        sb.AppendLine($"StartLocal: {(start.HasValue ? start.Value.ToString("O") : "null")}");
        sb.AppendLine($"EndLocal: {(end.HasValue ? end.Value.ToString("O") : "null")}");
        sb.AppendLine($"DurationMinutes: {(start.HasValue && end.HasValue ? (end.Value - start.Value).TotalMinutes.ToString("0.##") : "null")}");
        sb.AppendLine($"Exception: {ex}");
        sb.AppendLine($"HResult: 0x{ex.HResult:X8}");

        var inner = ex.InnerException;
        var depth = 0;
        while (inner != null)
        {
            sb.AppendLine($"Inner[{depth}] Type={inner.GetType().FullName} HResult=0x{inner.HResult:X8} Message={inner.Message}");
            sb.AppendLine(inner.ToString());
            inner = inner.InnerException;
            depth++;
        }

        return sb.ToString();
    }
}
