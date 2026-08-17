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
    private const int OlMailItem = 0;
    private const int OlFolderCalendar = 9;
    private const int MaxCalendarFetchIterations = 10000;
    private const int MaxHomeOfficeSearchIterations = 64;
    private const int MaxRepeatedCalendarItemCount = 25;
    private const int OlBusy = 2;
    // Microsoft.Office.Interop.Outlook.OlBusyStatus.olWorkingElsewhere.
    private const int OlWorkingElsewhere = 4;
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

    public (bool ok, string error) DeleteBlock(string? entryId, bool ignoreSyncDisabled = false)
    {
        if ((!_settings.Current.OutlookSyncEnabled && !ignoreSyncDisabled) || string.IsNullOrWhiteSpace(entryId))
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

    public (bool ok, string entryId, string error) UpsertHomeOfficeAppointment(string? existingEntryId, DateTime day)
    {
        try
        {
            return ExecuteOnSta<(bool ok, string entryId, string error)>(() =>
            {
                var outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null) return (false, string.Empty, "Outlook ist nicht installiert.");
                object? app = null; object? ns = null; object? item = null;
                try
                {
                    app = CreateOrAttachOutlook(outlookType);
                    if (app == null) return (false, string.Empty, "Outlook konnte nicht gestartet werden.");
                    dynamic appDyn = app;
                    ns = appDyn.GetNamespace("MAPI");
                    TryLogon(ns);
                    dynamic nsDyn = ns!;
                    item = string.IsNullOrWhiteSpace(existingEntryId)
                        ? appDyn.CreateItem(OlAppointmentItem)
                        : nsDyn.GetItemFromID(existingEntryId);
                    if (item == null) return (false, string.Empty, "Outlook-Termin konnte nicht erstellt werden.");
                    dynamic appointment = item;
                    var title = GetPlenaroHomeOfficeTitle(day);
                    appointment.Subject = title;
                    appointment.Body = title;
                    appointment.Start = day.Date;
                    appointment.End = day.Date.AddDays(1);
                    appointment.AllDayEvent = true;
                    appointment.ReminderSet = false;
                    appointment.BusyStatus = OlWorkingElsewhere;
                    appointment.Location = "An anderem Ort tätig";
                    appointment.Categories = string.Empty;
                    appointment.Save();

                    var savedSubject = Convert.ToString(appointment.Subject) ?? string.Empty;
                    var savedStart = Convert.ToDateTime(appointment.Start);
                    var savedEnd = Convert.ToDateTime(appointment.End);
                    var savedAllDay = Convert.ToBoolean(appointment.AllDayEvent);
                    var savedBusyStatus = Convert.ToInt32(appointment.BusyStatus);
                    var savedLocation = Convert.ToString(appointment.Location) ?? string.Empty;
                    var savedCategories = Convert.ToString(appointment.Categories) ?? string.Empty;
                    var savedBodyLength = (Convert.ToString(appointment.Body) ?? string.Empty).Length;
                    var savedEntryId = Convert.ToString(appointment.EntryID) ?? string.Empty;
                    _logger.Info($"[HomeOfficeOutlookSaved] date={day:yyyy-MM-dd} subject='{savedSubject}' start={savedStart:O} end={savedEnd:O} allDay={savedAllDay.ToString().ToLowerInvariant()} busyStatus={savedBusyStatus} location='{savedLocation}' categories='{savedCategories}' bodyLength={savedBodyLength} entryId='{savedEntryId}'");
                    return (true, savedEntryId, string.Empty);
                }
                finally
                {
                    SafeReleaseComObject(item); SafeReleaseComObject(ns); SafeReleaseComObject(app);
                }
            });
        }
        catch (Exception ex)
        {
            _logger.Error(BuildOutlookExceptionLog("UpsertHomeOfficeAppointment", ex, day.Date, day.Date.AddDays(1)));
            return (false, string.Empty, BuildUserFacingOutlookError(ex));
        }
    }

    public (bool ok, List<OutlookCalendarEvent> events, string error) FindPlenaroHomeOfficeAppointmentsForDay(DateTime day)
    {
        day = day.Date;
        try
        {
            return ExecuteOnSta<(bool ok, List<OutlookCalendarEvent> events, string error)>(() =>
            {
                var outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null)
                    return (false, new List<OutlookCalendarEvent>(), "Outlook ist nicht installiert.");

                object? app = null; object? ns = null; object? folder = null; object? items = null; object? restricted = null;
                try
                {
                    app = CreateOrAttachOutlook(outlookType);
                    if (app == null) return (false, new List<OutlookCalendarEvent>(), "Outlook konnte nicht gestartet werden.");
                    dynamic appDyn = app;
                    ns = appDyn.GetNamespace("MAPI");
                    TryLogon(ns);
                    dynamic nsDyn = ns!;
                    folder = nsDyn.GetDefaultFolder(OlFolderCalendar);
                    dynamic folderDyn = folder!;
                    var calendarName = Convert.ToString(folderDyn.Name) ?? string.Empty;
                    items = folderDyn.Items;
                    dynamic itemsDyn = items!;
                    itemsDyn.Sort("[Start]", false);
                    itemsDyn.IncludeRecurrences = true;

                    var culture = CultureInfo.CurrentCulture;
                    var fromFilter = FormatOutlookRestrictDate(day, culture);
                    var toFilter = FormatOutlookRestrictDate(day.AddDays(1), culture);
                    var title = GetPlenaroHomeOfficeTitle(day);
                    var filter = $"[Start] < '{toFilter}' AND [End] > '{fromFilter}' AND [Subject] = '{title}'";
                    restricted = itemsDyn.Restrict(filter);

                    var matches = new List<OutlookCalendarEvent>();
                    dynamic collection = restricted!;
                    object? item = collection.GetFirst();
                    var scanned = 0;
                    while (item != null && scanned < MaxHomeOfficeSearchIterations)
                    {
                        try
                        {
                            scanned++;
                            if (TryReadCalendarEvent(item, calendarName, day, day.AddDays(1), out var calendarEvent, out _)
                                && calendarEvent != null
                                && HasPlenaroHomeOfficeSignature(calendarEvent, day))
                                matches.Add(calendarEvent);
                        }
                        finally
                        {
                            SafeReleaseComObject(item);
                        }
                        item = collection.GetNext();
                    }
                    var limitReached = item != null;
                    if (limitReached) SafeReleaseComObject(item);
                    _logger.Info($"[HomeOfficeOutlookSearch] date={day:yyyy-MM-dd} scanned={scanned} matched={matches.Count} bounded=true fallback=false");
                    if (limitReached)
                        return (false, new List<OutlookCalendarEvent>(), "Die begrenzte Homeoffice-Suche hat ihr Sicherheitslimit erreicht.");
                    return (true, matches, string.Empty);
                }
                finally
                {
                    SafeReleaseComObject(restricted); SafeReleaseComObject(items); SafeReleaseComObject(folder); SafeReleaseComObject(ns); SafeReleaseComObject(app);
                }
            });
        }
        catch (Exception ex)
        {
            _logger.Error(BuildOutlookExceptionLog("FindPlenaroHomeOfficeAppointmentsForDay", ex, day, day.AddDays(1)));
            return (false, new List<OutlookCalendarEvent>(), BuildUserFacingOutlookError(ex));
        }
    }

    private static string GetPlenaroHomeOfficeTitle(DateTime day) => $"Homeoffice am {day:yyyy-MM-dd}";

    private static bool HasPlenaroHomeOfficeSignature(OutlookCalendarEvent appointment, DateTime day)
    {
        var title = GetPlenaroHomeOfficeTitle(day);
        return appointment.IsAllDay
            && appointment.StartLocal.Date == day.Date
            && appointment.EndLocal == day.Date.AddDays(1)
            && string.Equals(appointment.Subject, title, StringComparison.Ordinal)
            && string.Equals(appointment.BodyPreview, title, StringComparison.Ordinal)
            && string.Equals(appointment.Location, "An anderem Ort tätig", StringComparison.OrdinalIgnoreCase)
            && int.TryParse(appointment.BusyStatus, NumberStyles.Integer, CultureInfo.InvariantCulture, out var busyStatus)
            && busyStatus == OlWorkingElsewhere;
    }

    public (bool ok, string error) SendHomeOfficeMail(DateTime day, IReadOnlyCollection<string> recipients)
    {
        try
        {
            return ExecuteOnSta<(bool ok, string error)>(() =>
            {
                var outlookType = Type.GetTypeFromProgID("Outlook.Application");
                if (outlookType == null) return (false, "Outlook ist nicht installiert.");
                object? app = null; object? mail = null; object? recipient = null;
                try
                {
                    app = CreateOrAttachOutlook(outlookType);
                    if (app == null) return (false, "Outlook konnte nicht gestartet werden.");
                    dynamic appDyn = app;
                    mail = appDyn.CreateItem(OlMailItem);
                    dynamic mailDyn = mail!;
                    foreach (var address in recipients)
                    {
                        recipient = mailDyn.Recipients.Add(address);
                        SafeReleaseComObject(recipient);
                        recipient = null;
                    }
                    if (!mailDyn.Recipients.ResolveAll()) return (false, "Mindestens ein Homeoffice-Empfänger konnte nicht aufgelöst werden.");
                    mailDyn.Subject = $"Homeoffice am {day:dd.MM.yyyy}";
                    mailDyn.Body = $"Hallo,{Environment.NewLine}{Environment.NewLine}ich mache am {day:dd.MM.yyyy} Homeoffice.";
                    mailDyn.Send();
                    return (true, string.Empty);
                }
                finally
                {
                    SafeReleaseComObject(recipient); SafeReleaseComObject(mail); SafeReleaseComObject(app);
                }
            });
        }
        catch (Exception ex)
        {
            _logger.Error(BuildOutlookExceptionLog("SendHomeOfficeMail", ex, day.Date, day.Date.AddDays(1)));
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

    public (bool ok, List<OutlookCalendarEvent> events, string error) GetCalendarEvents(DateTime fromLocal, DateTime toLocal, bool ignoreCalendarDisabled = false)
    {
        if (!_settings.Current.OutlookCalendarEnabled && !ignoreCalendarDisabled)
            return (true, new List<OutlookCalendarEvent>(), string.Empty);

        if (toLocal <= fromLocal)
            return (false, new List<OutlookCalendarEvent>(), "Ungültiger Zeitraum für Kalenderabfrage.");

        try
        {
            return ExecuteOnSta<(bool ok, List<OutlookCalendarEvent> events, string error)>(() =>
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
                    var rawItemsCount = ReadComCollectionCount(items);
                    itemsDyn.Sort("[Start]", false);
                    itemsDyn.IncludeRecurrences = true;

                    var normalizedFrom = fromLocal.Date;
                    var normalizedTo = toLocal.Date;
                    var outlookCulture = CultureInfo.CurrentCulture;
                    var fromFilter = FormatOutlookRestrictDate(normalizedFrom, outlookCulture);
                    var toFilter = FormatOutlookRestrictDate(normalizedTo, outlookCulture);
                    var filter = $"[Start] < '{toFilter}' AND [End] > '{fromFilter}'";
                    _logger.Info($"[OutlookFetchRestrict] culture='{outlookCulture.Name}' fromInclusive={normalizedFrom:O} toExclusive={normalizedTo:O} filter='{filter}' rawItemsCount={rawItemsCount} restrictedCount=<pending> includeRecurrences=True");

                    Exception? restrictException = null;
                    try
                    {
                        restricted = itemsDyn.Restrict(filter);
                    }
                    catch (Exception ex)
                    {
                        restrictException = ex;
                        _logger.Error($"[OutlookFetchRestrict] RestrictFailed error='{ex.Message}' filter='{filter}'");
                    }

                    if (restricted != null)
                    {
                        var restrictedCount = ReadComCollectionCount(restricted);
                        var restrictedResult = CollectCalendarEventsBounded(restricted, calendarName, normalizedFrom, normalizedTo, "Restrict");
                        _logger.Info($"[OutlookFetchRestrict] culture='{outlookCulture.Name}' filter='{filter}' rawItemsCount={rawItemsCount} restrictedCount={restrictedCount} includeRecurrences=True scanned={restrictedResult.Scanned} matched={restrictedResult.Events.Count} abortReason='{restrictedResult.AbortReason}'");
                        if (restrictedResult.Events.Count > 0 && restrictedResult.Completed)
                            return (true, restrictedResult.Events, string.Empty);
                    }

                    var fallbackResult = CollectCalendarEventsBounded(items, calendarName, normalizedFrom, normalizedTo, "Fallback");
                    _logger.Info($"[OutlookFetchFallback] scanned={fallbackResult.Scanned} matched={fallbackResult.Events.Count} abortReason='{fallbackResult.AbortReason}'");
                    if (fallbackResult.Completed)
                        return (true, fallbackResult.Events, string.Empty);

                    var fallbackError = restrictException == null
                        ? $"Outlook Kalender-Fallback wurde unvollständig beendet ({fallbackResult.AbortReason})."
                        : $"Outlook Kalenderfilter fehlgeschlagen ({restrictException.Message}); Fallback wurde unvollständig beendet ({fallbackResult.AbortReason}).";
                    return (false, new List<OutlookCalendarEvent>(), fallbackError);
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


    private CalendarCollectionResult CollectCalendarEventsBounded(object source, string calendarName, DateTime fromInclusive, DateTime toExclusive, string sourceName)
    {
        var events = new List<OutlookCalendarEvent>();
        var scanned = 0;
        var repeatedItemCount = 0;
        var previousFingerprint = string.Empty;
        var abortReason = "EndOfCollection";
        var completed = true;
        dynamic collection = source;
        object? itemObj;
        try
        {
            itemObj = collection.GetFirst();
        }
        catch (Exception ex)
        {
            return new CalendarCollectionResult(events, scanned, false, $"GetFirstFailed:{ex.Message}");
        }

        while (itemObj != null)
        {
            try
            {
                scanned++;
                if (scanned > MaxCalendarFetchIterations)
                {
                    abortReason = "MaxIterations";
                    completed = false;
                    break;
                }

                var itemStart = ReadComDate(itemObj, "Start");
                if (itemStart.HasValue && itemStart.Value >= toExclusive)
                {
                    abortReason = "StartOutsideTargetRange";
                    break;
                }

                var fingerprint = $"{ReadComString(itemObj, "EntryID")}|{FormatNullableDate(itemStart)}|{FormatNullableDate(ReadComDate(itemObj, "End"))}";
                if (string.Equals(previousFingerprint, fingerprint, StringComparison.Ordinal))
                    repeatedItemCount++;
                else
                    repeatedItemCount = 0;
                previousFingerprint = fingerprint;
                if (repeatedItemCount >= MaxRepeatedCalendarItemCount)
                {
                    abortReason = "RepeatedItemGuard";
                    completed = false;
                    break;
                }

                LogRawItem(itemObj, sourceName);

                if (!TryReadCalendarEvent(itemObj, calendarName, fromInclusive, toExclusive, out OutlookCalendarEvent? evt, out var rejectReason))
                {
                    LogRejectedItem(itemObj, rejectReason);
                }
                else
                {
                    _logger.Info($"[OutlookItemAccepted] subject='{evt!.Subject}' start={evt.StartLocal:O} end={evt.EndLocal:O} whyAccepted=CalendarItemWithValidRangeAndOverlap entryId='{evt.EntryId}' source='{sourceName}'");
                    events.Add(evt);
                }
            }
            finally
            {
                SafeReleaseComObject(itemObj);
            }

            try
            {
                itemObj = collection.GetNext();
            }
            catch (Exception ex)
            {
                abortReason = $"GetNextFailed:{ex.Message}";
                completed = false;
                itemObj = null;
            }
        }

        return new CalendarCollectionResult(events, scanned, completed, abortReason);
    }

    private sealed record CalendarCollectionResult(List<OutlookCalendarEvent> Events, int Scanned, bool Completed, string AbortReason);

    private void LogRawItem(object? rawItem, string sourceName)
    {
        var runtimeType = rawItem?.GetType().FullName ?? "<null>";
        var messageClass = ReadComString(rawItem, "MessageClass");
        var subject = ReadComString(rawItem, "Subject");
        var start = ReadComDate(rawItem, "Start");
        var end = ReadComDate(rawItem, "End");
        var allDay = ReadComBool(rawItem, "AllDayEvent");
        var meetingStatus = ReadComString(rawItem, "MeetingStatus");
        var busyStatus = ReadComString(rawItem, "BusyStatus");
        var isRecurring = ReadComBool(rawItem, "IsRecurring");
        var entryId = ReadComString(rawItem, "EntryID");

        _logger.Info($"[OutlookRawItem] source='{sourceName}' runtimeType='{runtimeType}' messageClass='{messageClass}' subject='{subject}' start={FormatNullableDate(start)} end={FormatNullableDate(end)} allDay={FormatNullableBool(allDay)} meetingStatus='{meetingStatus}' busyStatus='{busyStatus}' isRecurring={FormatNullableBool(isRecurring)} entryId='{entryId}'");
    }

    private void LogRejectedItem(object? rawItem, string reason)
    {
        var runtimeType = rawItem?.GetType().FullName ?? "<null>";
        var messageClass = ReadComString(rawItem, "MessageClass");
        var subject = ReadComString(rawItem, "Subject");

        _logger.Info($"[OutlookItemRejected] subject='{(string.IsNullOrWhiteSpace(subject) ? "<null>" : subject)}' runtimeType='{runtimeType}' messageClass='{messageClass}' reason={reason}");
    }

    private bool TryReadCalendarEvent(object? rawItem, string calendarName, DateTime fromInclusive, DateTime toExclusive, out OutlookCalendarEvent? calendarEvent, out string rejectReason)
    {
        calendarEvent = null;
        rejectReason = string.Empty;

        if (rawItem == null)
        {
            rejectReason = "NullItem";
            return false;
        }

        dynamic item = rawItem;

        var runtimeType = rawItem.GetType().FullName ?? rawItem.GetType().Name;
        var messageClass = ReadComString(rawItem, "MessageClass");
        var entryId = ReadComString(rawItem, "EntryID");
        var subject = ReadComString(rawItem, "Subject");

        DateTime start;
        DateTime end;
        try
        {
            start = NormalizeOutlookDateTime(Convert.ToDateTime(item.Start));
            end = NormalizeOutlookDateTime(Convert.ToDateTime(item.End));
        }
        catch
        {
            rejectReason = "MissingOrInvalidStartEnd";
            return false;
        }

        if (!IsCalendarLikeItem(messageClass, runtimeType, hasStartEnd: true))
        {
            rejectReason = "FilteredByItemType";
            return false;
        }

        var overlap = start < toExclusive && end > fromInclusive;
        if (!overlap)
        {
            rejectReason = "FilteredByTimeRange";
            return false;
        }

        var body = ReadComString(rawItem, "Body");
        var location = ReadComString(rawItem, "Location");
        var joinUrl = ExtractTeamsUrl(body, location);
        var busyStatus = ReadComString(rawItem, "BusyStatus");
        var sensitivity = ReadComString(rawItem, "Sensitivity");
        var categories = ReadComString(rawItem, "Categories");
        var organizer = ReadComString(rawItem, "Organizer");
        var iCalUid = ReadComString(rawItem, "GlobalAppointmentID");
        var meetingStatus = ReadComString(rawItem, "MeetingStatus");

        var allDay = ReadComBool(rawItem, "AllDayEvent") ?? false;
        var isPrivate = ReadComBool(rawItem, "IsPrivate") ?? false;
        var isRecurring = ReadComBool(rawItem, "IsRecurring") ?? false;
        var isCancelled = ReadComBool(rawItem, "IsCancelled") ?? false;

        var recurrenceState = ReadComInt(rawItem, "RecurrenceState") ?? 0;
        var isInstance = recurrenceState == 2 || recurrenceState == 3;

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
            BodyPreview = body.Length > 240 ? body.Substring(0, 240) : body,
            OnlineMeetingJoinUrl = joinUrl,
            Categories = categories
        };

        return true;
    }

    private static bool IsCalendarLikeItem(string messageClass, string runtimeType, bool hasStartEnd)
    {
        if (!string.IsNullOrWhiteSpace(messageClass))
        {
            if (messageClass.StartsWith("IPM.Appointment", StringComparison.OrdinalIgnoreCase))
                return true;


            if (messageClass.StartsWith("IPM.Task", StringComparison.OrdinalIgnoreCase)
                || messageClass.StartsWith("IPM.Note", StringComparison.OrdinalIgnoreCase)
                || messageClass.StartsWith("IPM.StickyNote", StringComparison.OrdinalIgnoreCase))
                return false;
        }

        if (runtimeType.Contains("Appointment", StringComparison.OrdinalIgnoreCase))
            return true;

        return hasStartEnd;
    }

    private static string ReadComString(object? rawItem, string propertyName)
    {
        var value = ReadComObject(rawItem, propertyName);
        return value == null ? string.Empty : Convert.ToString(value) ?? string.Empty;
    }

    private static DateTime? ReadComDate(object? rawItem, string propertyName)
    {
        var value = ReadComObject(rawItem, propertyName);
        if (value == null)
            return null;

        try
        {
            return NormalizeOutlookDateTime(Convert.ToDateTime(value));
        }
        catch
        {
            return null;
        }
    }

    private static bool? ReadComBool(object? rawItem, string propertyName)
    {
        var value = ReadComObject(rawItem, propertyName);
        if (value == null)
            return null;

        try
        {
            return Convert.ToBoolean(value);
        }
        catch
        {
            return null;
        }
    }

    private static int? ReadComInt(object? rawItem, string propertyName)
    {
        var value = ReadComObject(rawItem, propertyName);
        if (value == null)
            return null;

        try
        {
            return Convert.ToInt32(value);
        }
        catch
        {
            return null;
        }
    }

    private static object? ReadComObject(object? rawItem, string propertyName)
    {
        if (rawItem == null)
            return null;

        dynamic item = rawItem;
        return propertyName switch
        {
            "MessageClass" => SafeRead(() => (object?)item.MessageClass),
            "Subject" => SafeRead(() => (object?)item.Subject),
            "Body" => SafeRead(() => (object?)item.Body),
            "Location" => SafeRead(() => (object?)item.Location),
            "Sensitivity" => SafeRead(() => (object?)item.Sensitivity),
            "Categories" => SafeRead(() => (object?)item.Categories),
            "Organizer" => SafeRead(() => (object?)item.Organizer),
            "GlobalAppointmentID" => SafeRead(() => (object?)item.GlobalAppointmentID),
            "Start" => SafeRead(() => (object?)item.Start),
            "End" => SafeRead(() => (object?)item.End),
            "AllDayEvent" => SafeRead(() => (object?)item.AllDayEvent),
            "MeetingStatus" => SafeRead(() => (object?)item.MeetingStatus),
            "BusyStatus" => SafeRead(() => (object?)item.BusyStatus),
            "IsRecurring" => SafeRead(() => (object?)item.IsRecurring),
            "IsPrivate" => SafeRead(() => (object?)item.IsPrivate),
            "IsCancelled" => SafeRead(() => (object?)item.IsCancelled),
            "RecurrenceState" => SafeRead(() => (object?)item.RecurrenceState),
            "EntryID" => SafeRead(() => (object?)item.EntryID),
            _ => null
        };
    }

    private static string FormatNullableDate(DateTime? value)
        => value.HasValue ? value.Value.ToString("O") : "<null>";

    private static string FormatNullableBool(bool? value)
        => value.HasValue ? value.Value.ToString() : "<null>";

    private static int ReadComCollectionCount(object? collection)
    {
        if (collection == null)
            return -1;

        try
        {
            dynamic value = collection;
            return Convert.ToInt32(value.Count);
        }
        catch
        {
            return -1;
        }
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

    private static string FormatOutlookRestrictDate(DateTime value, CultureInfo culture)
    {
        var local = value.Kind switch
        {
            DateTimeKind.Utc => value.ToLocalTime(),
            DateTimeKind.Local => value,
            _ => DateTime.SpecifyKind(value, DateTimeKind.Local)
        };
        return local.ToString("g", culture);
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
