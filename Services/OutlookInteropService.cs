using System.Collections.Generic;
using System.Globalization;
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
    private const int SW_RESTORE = 9;

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



    public (bool ok, string error) OpenCalendarEvent(string entryId)
    {
        if (string.IsNullOrWhiteSpace(entryId))
            return (false, "Outlook EntryID fehlt.");

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
                object? inspector = null;

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
                        return (false, "Outlook Termin nicht gefunden.");

                    dynamic itemDyn = item;
                    itemDyn.Display(false);

                    inspector = itemDyn.GetInspector;
                    dynamic inspDyn = inspector!;
                    inspDyn.Display();
                    inspDyn.Activate();

                    IntPtr hwnd = IntPtr.Zero;
                    try
                    {
                        hwnd = new IntPtr(Convert.ToInt32(inspDyn.Hwnd));
                    }
                    catch
                    {
                        try { hwnd = new IntPtr(Convert.ToInt32(inspDyn.WindowHandle)); } catch { }
                    }

                    if (hwnd != IntPtr.Zero)
                    {
                        ShowWindow(hwnd, SW_RESTORE);
                        SetForegroundWindow(hwnd);
                    }
                    else
                    {
                        try { appDyn.ActiveExplorer()?.Activate(); } catch { }
                        try { inspDyn.Activate(); } catch { }
                    }

                    return (true, string.Empty);
                }
                finally
                {
                    SafeReleaseComObject(inspector);
                    SafeReleaseComObject(item);
                    SafeReleaseComObject(ns);
                    SafeReleaseComObject(app);
                }
            });
        }
        catch (Exception ex)
        {
            _logger.Error(BuildOutlookExceptionLog("OpenCalendarEvent", ex, null, null));
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
                    var calendarName = Convert.ToString(folderDyn.Name) ?? string.Empty;
                    var folderEntryId = Convert.ToString(folderDyn.EntryID) ?? string.Empty;
                    var storeId = Convert.ToString(folderDyn.StoreID) ?? string.Empty;
                    string storeName;
                    try { storeName = Convert.ToString(folderDyn.Store?.DisplayName) ?? string.Empty; } catch { storeName = string.Empty; }
                    _logger.Info($"[OutlookFetchFolder] calendarName='{calendarName}' folderEntryId='{folderEntryId}' storeId='{storeId}' storeName='{storeName}'");

                    items = folderDyn.Items;
                    dynamic itemsDyn = items!;
                    itemsDyn.IncludeRecurrences = true;
                    itemsDyn.Sort("[Start]");

                    var fromFilter = FormatOutlookRestrictDate(fromLocal);
                    var toFilter = FormatOutlookRestrictDate(toLocal);
                    var filter = $"[Start] < '{toFilter}' AND [End] > '{fromFilter}'";
                    _logger.Info($"[OutlookFetchRestrict] fromInclusive={fromLocal:O} toExclusive={toLocal:O} filter='{filter}'");
                    restricted = itemsDyn.Restrict(filter);

                    LogProbeScanForMissingDays(itemsDyn, fromLocal, toLocal);

                    var events = new List<OutlookCalendarEvent>();
                    foreach (var raw in (System.Collections.IEnumerable)restricted!)
                    {
                        object? appointment = raw;
                        try
                        {
                            dynamic a = appointment!;
                            var startRaw = Convert.ToDateTime(a.Start);
                            var endRaw = Convert.ToDateTime(a.End);
                            DateTime start = NormalizeOutlookDateTime(startRaw);
                            DateTime end = NormalizeOutlookDateTime(endRaw);
                            if (end <= fromLocal || start >= toLocal)
                                continue;

                            var body = Convert.ToString(a.Body) ?? string.Empty;
                            var location = Convert.ToString(a.Location) ?? string.Empty;
                            var joinUrl = ExtractTeamsUrl(body, location);

                            var entryId = Convert.ToString(a.EntryID) ?? string.Empty;
                            var busyRaw = Convert.ToString(a.BusyStatus) ?? string.Empty;
                            var iCalUid = Convert.ToString(a.GlobalAppointmentID) ?? string.Empty;
                            var allDayEvent = Convert.ToBoolean(a.AllDayEvent);
                            var subjectRaw = Convert.ToString(a.Subject) ?? string.Empty;
                            _logger.Info($"[OutlookRawEvent] subject='{subjectRaw}' start={start:O} end={end:O} allDay={allDayEvent} startKind={start.Kind} endKind={end.Kind} busyStatus='{busyRaw}' sensitivityRaw='{Convert.ToString(a.Sensitivity) ?? string.Empty}' isRecurringRaw='{Convert.ToString(a.IsRecurring) ?? string.Empty}' entryId='{entryId}' calendar='{calendarName}'");

                            string sensitivity;
                            bool isPrivate;
                            bool isRecurring;
                            bool isInstance;
                            try { sensitivity = Convert.ToString(a.Sensitivity) ?? string.Empty; } catch { sensitivity = string.Empty; }
                            try { isPrivate = Convert.ToInt32(a.Sensitivity) == 2; } catch { try { isPrivate = Convert.ToBoolean(a.IsPrivate); } catch { isPrivate = false; } }
                            try { isRecurring = Convert.ToBoolean(a.IsRecurring); } catch { isRecurring = false; }
                            try { isInstance = Convert.ToInt32(a.RecurrenceState) == 2 || Convert.ToInt32(a.RecurrenceState) == 3; } catch { isInstance = false; }

                            events.Add(new OutlookCalendarEvent
                            {
                                Id = string.IsNullOrWhiteSpace(entryId) ? Guid.NewGuid().ToString("N") : entryId,
                                EntryId = entryId,
                                ICalUId = iCalUid,
                                CalendarName = calendarName,
                                BusyStatus = busyRaw,
                                Sensitivity = sensitivity,
                                IsPrivate = isPrivate,
                                IsRecurring = isRecurring,
                                IsInstance = isInstance,
                                Subject = string.IsNullOrWhiteSpace(subjectRaw) ? "(Kein Betreff)" : subjectRaw,
                                StartLocal = start,
                                EndLocal = end,
                                IsAllDay = allDayEvent,
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


    private static DateTime NormalizeOutlookDateTime(DateTime value)
    {
        if (value.Kind == DateTimeKind.Utc)
            return value.ToLocalTime();

        if (value.Kind == DateTimeKind.Unspecified)
            return DateTime.SpecifyKind(value, DateTimeKind.Local);

        return value;
    }

    private static string FormatOutlookRestrictDate(DateTime value)
    {
        var local = value.Kind == DateTimeKind.Local ? value : value.ToLocalTime();
        return local.ToString("MM/dd/yyyy hh:mm tt", CultureInfo.GetCultureInfo("en-US"));
    }

    private void LogProbeScanForMissingDays(dynamic itemsDyn, DateTime fromInclusive, DateTime toExclusive)
    {
        var probeDays = new[]
        {
            new DateTime(2026, 3, 4),
            new DateTime(2026, 3, 6)
        };

        _logger.Info($"[OutlookProbeDayScan] days=2026-03-04,2026-03-06 mode=IterateAllItemsNoRestrict fromInclusive={fromInclusive:O} toExclusive={toExclusive:O}");

        object? raw = null;
        try
        {
            foreach (var item in (System.Collections.IEnumerable)itemsDyn)
            {
                raw = item;
                try
                {
                    dynamic a = item;
                    var start = NormalizeOutlookDateTime(Convert.ToDateTime(a.Start));
                    var end = NormalizeOutlookDateTime(Convert.ToDateTime(a.End));
                    var subject = Convert.ToString(a.Subject) ?? string.Empty;
                    var allDay = Convert.ToBoolean(a.AllDayEvent);
                    var entryId = Convert.ToString(a.EntryID) ?? string.Empty;

                    foreach (var day in probeDays)
                    {
                        var dayStart = day.Date;
                        var dayEnd = dayStart.AddDays(1);
                        var overlap = start < dayEnd && end > dayStart;
                        if (!overlap)
                            continue;

                        _logger.Info($"[OutlookProbeDayHit] day={day:yyyy-MM-dd} subject='{subject}' start={start:O} end={end:O} allDay={allDay} entryId='{entryId}' inRequestedRange={start < toExclusive && end > fromInclusive}");
                    }
                }
                catch
                {
                    // ignore probe conversion issues
                }
                finally
                {
                    SafeReleaseComObject(raw);
                    raw = null;
                }
            }
        }
        catch
        {
            // probe scan is diagnostic only
        }
        finally
        {
            SafeReleaseComObject(raw);
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


    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

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
