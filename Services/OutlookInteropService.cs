using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Linq;
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
                    _logger.Info($"[OutlookFetchFolder] folderName='{calendarName}' folderEntryId='{folderEntryId}' storeId='{storeId}' storeName='{storeName}'");

                    items = folderDyn.Items;
                    dynamic itemsDyn = items!;
                    itemsDyn.IncludeRecurrences = true;
                    itemsDyn.Sort("[Start]");

                    var normalizedFrom = fromLocal.Date;
                    var normalizedTo = toLocal.Date;
                    var fromFilter = FormatOutlookRestrictDate(normalizedFrom);
                    var toFilter = FormatOutlookRestrictDate(normalizedTo);
                    var filter = $"[Start] < '{toFilter}' AND [End] > '{fromFilter}'";
                    _logger.Info($"[OutlookFetchRestrict] fromInclusive={normalizedFrom:O} toExclusive={normalizedTo:O} filter='{filter}'");

                    try
                    {
                        restricted = itemsDyn.Restrict(filter);
                    }
                    catch (Exception ex)
                    {
                        _logger.Error($"[OutlookFetchRestrict] RestrictFailed error='{ex.Message}' filter='{filter}'");
                        restricted = items;
                    }

                    LogProbeScanForMissingDays(itemsDyn, normalizedFrom, normalizedTo);

                    var events = new List<OutlookCalendarEvent>();
                    foreach (var raw in (System.Collections.IEnumerable)restricted!)
                    {
                        object? appointment = raw;
                        try
                        {
                            if (!TryReadCalendarEvent(appointment, calendarName, normalizedFrom, normalizedTo, out var evt))
                                continue;

                            events.Add(evt!);
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


    private bool TryReadCalendarEvent(object? rawItem, string calendarName, DateTime fromInclusive, DateTime toExclusive, out OutlookCalendarEvent? calendarEvent)
    {
        calendarEvent = null;
        if (rawItem == null)
            return false;

        dynamic item = rawItem;

        string messageClass;
        try { messageClass = Convert.ToString(item.MessageClass) ?? string.Empty; } catch { messageClass = string.Empty; }

        var className = rawItem.GetType().Name;
        var entryId = SafeRead(() => Convert.ToString(item.EntryID)) ?? string.Empty;
        var subject = SafeRead(() => Convert.ToString(item.Subject)) ?? string.Empty;

        if (!IsCalendarLikeItem(messageClass, className))
        {
            _logger.Info($"[OutlookEventFiltered] subject='{subject}' reason=FilteredByItemType className='{className}' messageClass='{messageClass}' entryId='{entryId}'");
            return false;
        }

        DateTime start;
        DateTime end;
        try
        {
            start = NormalizeOutlookDateTime(Convert.ToDateTime(item.Start));
            end = NormalizeOutlookDateTime(Convert.ToDateTime(item.End));
        }
        catch
        {
            _logger.Info($"[OutlookEventFiltered] subject='{subject}' reason=MissingOrInvalidStartEnd className='{className}' messageClass='{messageClass}' entryId='{entryId}'");
            return false;
        }

        var overlap = start < toExclusive && end > fromInclusive;
        if (!overlap)
        {
            _logger.Info($"[OutlookEventFiltered] subject='{subject}' reason=FilteredByTimeRange start={start:O} end={end:O} fromInclusive={fromInclusive:O} toExclusive={toExclusive:O} entryId='{entryId}'");
            return false;
        }

        var body = SafeRead(() => Convert.ToString(item.Body)) ?? string.Empty;
        var location = SafeRead(() => Convert.ToString(item.Location)) ?? string.Empty;
        var joinUrl = ExtractTeamsUrl(body, location);
        var busyStatus = SafeRead(() => Convert.ToString(item.BusyStatus)) ?? string.Empty;
        var sensitivity = SafeRead(() => Convert.ToString(item.Sensitivity)) ?? string.Empty;
        var categories = SafeRead(() => Convert.ToString(item.Categories)) ?? string.Empty;
        var organizer = SafeRead(() => Convert.ToString(item.Organizer)) ?? string.Empty;
        var iCalUid = SafeRead(() => Convert.ToString(item.GlobalAppointmentID)) ?? string.Empty;
        var meetingStatus = SafeRead(() => Convert.ToString(item.MeetingStatus)) ?? string.Empty;

        bool allDay = SafeRead(() => Convert.ToBoolean(item.AllDayEvent));
        bool isPrivate = SafeRead(() => Convert.ToBoolean(item.IsPrivate));
        bool isRecurring = SafeRead(() => Convert.ToBoolean(item.IsRecurring));
        bool isCancelled = SafeRead(() => Convert.ToBoolean(item.IsCancelled));

        var recurrenceState = SafeRead(() => Convert.ToInt32(item.RecurrenceState));
        var isInstance = recurrenceState == 2 || recurrenceState == 3;

        _logger.Info($"[OutlookRawEvent] subject='{subject}' start={start:O} end={end:O} isAllDay={allDay} busyStatus='{busyStatus}' sensitivity='{sensitivity}' isPrivate={isPrivate} isRecurring={isRecurring} isInstance={isInstance} meetingStatus='{meetingStatus}' messageClass='{messageClass}' location='{location}' categories='{categories}' entryId='{entryId}'");

        calendarEvent = new OutlookCalendarEvent
        {
            Id = string.IsNullOrWhiteSpace(entryId) ? Guid.NewGuid().ToString("N") : entryId,
            EntryId = entryId,
            ICalUId = iCalUid,
            CalendarName = calendarName,
            BusyStatus = busyStatus,
            Sensitivity = sensitivity,
            IsPrivate = isPrivate,
            IsRecurring = isRecurring,
            IsInstance = isInstance,
            IsCancelled = isCancelled,
            MeetingStatus = meetingStatus,
            MessageClass = messageClass,
            Subject = string.IsNullOrWhiteSpace(subject) ? "(Kein Betreff)" : subject,
            StartLocal = start,
            EndLocal = end,
            IsAllDay = allDay,
            Location = location,
            Organizer = organizer,
            BodyPreview = body.Length > 240 ? body[..240] : body,
            OnlineMeetingJoinUrl = joinUrl,
            Categories = categories
        };

        return true;
    }

    private static bool IsCalendarLikeItem(string messageClass, string className)
    {
        if (!string.IsNullOrWhiteSpace(messageClass) && messageClass.StartsWith("IPM.Appointment", StringComparison.OrdinalIgnoreCase))
            return true;

        return className.Contains("Appointment", StringComparison.OrdinalIgnoreCase);
    }

    private static T SafeRead<T>(Func<T> getter, T fallback = default!)
    {
        try
        {
            return getter();
        }
        catch
        {
            return fallback;
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
        var probeDays = Enumerable.Range(0, Math.Max(1, (toExclusive.Date - fromInclusive.Date).Days))
            .Select(offset => fromInclusive.Date.AddDays(offset))
            .Take(14)
            .ToArray();

        _logger.Info($"[OutlookProbeDayScan] days={string.Join(',', probeDays.Select(d => d.ToString("yyyy-MM-dd")))} mode=IterateAllItemsNoRestrict fromInclusive={fromInclusive:O} toExclusive={toExclusive:O}");

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
