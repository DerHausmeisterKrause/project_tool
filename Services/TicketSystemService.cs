using System.Net;
using System.Net.Http;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using TaskTool.Models;
using TaskStatus = TaskTool.Models.TaskStatus;

namespace TaskTool.Services;

public class TicketSystemService : IDisposable
{
    private readonly SettingsService _settings;
    private readonly TaskService _tasks;
    private readonly LoggerService _logger;
    private readonly HttpClient _client = new() { Timeout = TimeSpan.FromSeconds(45) };
    private readonly SemaphoreSlim _syncGate = new(1, 1);
    private readonly SemaphoreSlim _dynamicFieldOptionsGate = new(1, 1);
    private readonly System.Threading.Timer _timer;
    private IReadOnlyDictionary<string, IReadOnlyList<TicketFieldOption>> _dynamicFieldOptionsCache = new Dictionary<string, IReadOnlyList<TicketFieldOption>>(StringComparer.OrdinalIgnoreCase);
    private DateTime _dynamicFieldOptionsCacheExpiresUtc;
    private bool _dynamicFieldOptionsCacheValid;

    public string LastError { get; private set; } = string.Empty;
    public event Action? TasksChanged;

    public TicketSystemService(SettingsService settings, TaskService tasks, LoggerService logger)
    {
        _settings = settings;
        _tasks = tasks;
        _logger = logger;
        _timer = new System.Threading.Timer(async _ => await SyncAssignedTicketsAsync("timer"), null, Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
        HandleSettingsChanged();
    }

    public void HandleSettingsChanged()
    {
        var interval = Math.Clamp(_settings.Current.TicketSystemSyncIntervalMinutes, 1, 1440);
        _timer.Change(TimeSpan.FromMinutes(interval), TimeSpan.FromMinutes(interval));
    }

    public Task<(int created, int updated, int skipped)> ImportAssignedOpenTicketsAsync()
        => SyncAssignedTicketsAsync("manual");

    public async Task<TicketBookingContext> GetTicketBookingContextAsync(TaskItem task)
    {
        var ticketId = ExtractZnunyTicketIdFromTask(task);
        if (string.IsNullOrWhiteSpace(ticketId))
            throw new InvalidOperationException("Der ausgewählte Task besitzt keine eindeutige Znuny-TicketID.");

        var configError = ValidateConfiguration(requireAgentId: false);
        if (!string.IsNullOrWhiteSpace(configError))
            throw new InvalidOperationException(configError);

        var sessionId = await CreateSessionAsync();
        var ticket = await GetTicketAsync(ticketId, sessionId, HashSessionId(sessionId))
                     ?? throw new InvalidOperationException("TicketGet lieferte keine Ticketdaten.");
        var costField = _settings.Current.TicketSystemCostCenterFieldName;
        var orderField = _settings.Current.TicketSystemOrderFieldName;
        var optionFields = await GetDynamicFieldOptionsAsync(sessionId, forceRefresh: false);
        var costOptions = GetFieldOptions(optionFields, costField, _settings.Current.TicketSystemCostCenterOptions);
        var orderOptions = GetFieldOptions(optionFields, orderField, _settings.Current.TicketSystemOrderOptions);
        var costCenterValue = ticket.GetDynamicFieldValue(costField);
        var orderValue = ticket.GetDynamicFieldValue(orderField);
        _logger.Info($"[ZnunyTicketDynamicFields] ticketId={ticket.TicketID} requestedDynamicFields=true availableFields=[{string.Join(',', ticket.DynamicFieldValues.Keys.OrderBy(name => name, StringComparer.OrdinalIgnoreCase))}] costCenterValue='{costCenterValue}' orderValue='{orderValue}'");
        var fieldsMissing = !ticket.DynamicFieldValues.ContainsKey(costField) || !ticket.DynamicFieldValues.ContainsKey(orderField);
        var information = fieldsMissing ? "Kostenstelle/Auftrag konnten nicht aus OTRS geladen werden." : string.Empty;
        LogDynamicFieldSelection(ticket.TicketID, costField, costCenterValue, costOptions);
        LogDynamicFieldSelection(ticket.TicketID, orderField, orderValue, orderOptions);

        return new TicketBookingContext(
            ticket.TicketID,
            ticket.TicketNumber,
            costCenterValue,
            orderValue,
            costOptions,
            orderOptions,
            information);
    }

    public void InvalidateDynamicFieldOptionsCache()
    {
        _dynamicFieldOptionsCache = new Dictionary<string, IReadOnlyList<TicketFieldOption>>(StringComparer.OrdinalIgnoreCase);
        _dynamicFieldOptionsCacheExpiresUtc = DateTime.MinValue;
        _dynamicFieldOptionsCacheValid = false;
    }

    public async Task<TicketBookingResult> BookTimeAsync(
        TaskItem task,
        long sourceSeconds,
        string shortDescription,
        string costCenter,
        string order)
    {
        var ticketId = ExtractZnunyTicketIdFromTask(task);
        var ticketNumber = ExtractZnunyTicketNumberFromTask(task);
        if (string.IsNullOrWhiteSpace(ticketId))
            return new TicketBookingResult(false, false, "Der Task ist keinem eindeutigen Znuny-Ticket zugeordnet.");
        if (sourceSeconds <= 0)
            return new TicketBookingResult(false, false, "Es ist keine noch nicht gebuchte Zeit vorhanden.");

        var configError = ValidateConfiguration(requireAgentId: false);
        if (!string.IsNullOrWhiteSpace(configError))
            return new TicketBookingResult(false, false, configError);

        TicketTimeBooking? booking = null;
        var serverConfirmed = false;
        try
        {
            var sessionId = await CreateSessionAsync();
            var sessionHash = HashSessionId(sessionId);
            var pending = _tasks.GetPendingTicketTimeBooking(task.Id);
            if (pending != null)
            {
                var currentTicket = await GetTicketAsync(ticketId, sessionId, sessionHash);
                var reconciledArticleId = currentTicket?.FindArticleIdContaining(BookingMarker(pending.BookingId));
                if (!string.IsNullOrWhiteSpace(reconciledArticleId))
                {
                    _tasks.CompleteTicketTimeBooking(pending, reconciledArticleId);
                    _logger.Info($"[ZnunyTimeBooking] ticketId={ticketId} taskId={task.Id} bookingId={pending.BookingId} articleId={reconciledArticleId} action=reconciled");
                    return new TicketBookingResult(true, false, $"{pending.Minutes:0.##} Min. erfasst, {pending.BookedMinutes:0.##} Min. in OTRS gebucht.");
                }

                return new TicketBookingResult(false, true,
                    "Der Status der vorherigen Buchung ist noch unklar. Sie wurde nicht erneut gesendet, um eine Doppelbuchung zu verhindern. Bitte später erneut abgleichen.");
            }

            var minutes = decimal.Round(sourceSeconds / 60m, 2, MidpointRounding.AwayFromZero);
            var bookedMinutes = Math.Ceiling(sourceSeconds / 900m) * 15m;
            var timeUnit = bookedMinutes;
            booking = new TicketTimeBooking
            {
                TaskId = task.Id,
                TicketId = ticketId,
                TicketNumber = ticketNumber,
                Minutes = minutes,
                BookedMinutes = bookedMinutes,
                SourceSeconds = sourceSeconds,
                ShortDescription = string.IsNullOrWhiteSpace(shortDescription) ? "Zeitbuchung" : shortDescription.Trim(),
                CostCenter = costCenter ?? string.Empty,
                Order = order ?? string.Empty
            };
            _tasks.CreateTicketTimeBooking(booking);

            var payload = BuildTicketTimeBookingPayload(ticketId, sessionId, booking, timeUnit);
            var route = ResolveTicketUpdateRoute(ticketId);
            LogTicketUpdateRequest(route, ticketId, booking, timeUnit);
            using var request = new HttpRequestMessage(HttpMethod.Post, Combine(_settings.Current.TicketSystemApiUrl, route))
            {
                Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
            };
            _logger.Info($"[ZnunyTimeBooking] route={route} ticketId={ticketId} taskId={task.Id} bookingId={booking.BookingId} minutes={minutes:0.##} timeUnit={timeUnit:0.####} action=send");
            var response = await SendZnunyAsync(request, "TicketUpdateTimeBooking", "[ZnunyTicketUpdateResponse]");
            EnsureTicketUpdateResponseIsInterpretable(response);
            serverConfirmed = true;
            var articleId = ExtractFirstValueRecursive(response.Body, "ArticleID");
            _tasks.CompleteTicketTimeBooking(booking, articleId);
            _logger.Info($"[ZnunyTimeBooking] ticketId={ticketId} taskId={task.Id} bookingId={booking.BookingId} articleId={articleId} action=completed");
            return new TicketBookingResult(true, false, $"{minutes:0.##} Min. erfasst, auf {bookedMinutes:0.##} Min. aufgerundet und erfolgreich in OTRS gebucht.");
        }
        catch (ZnunyApiException ex)
        {
            var uncertainServerFailure = booking != null && (int)ex.StatusCode >= 500;
            if (booking != null && !uncertainServerFailure)
                _tasks.FailTicketTimeBooking(booking);
            LogZnunyError(ex);
            if (uncertainServerFailure)
                return new TicketBookingResult(false, true, "Znuny meldete einen Serverfehler nach dem Sendeversuch. Die Booking-ID wird vor einem weiteren Versuch abgeglichen; es erfolgt keine automatische Doppelbuchung.");
            return new TicketBookingResult(false, false, FormatApiError("Zeitbuchung fehlgeschlagen", ex));
        }
        catch (Exception ex) when (booking != null && (ex is HttpRequestException || ex is TaskCanceledException || ex is JsonException))
        {
            _logger.Error($"[ZnunyTimeBooking] ticketId={ticketId} taskId={task.Id} bookingId={booking.BookingId} action=pending-reconciliation message={ex.Message}");
            return new TicketBookingResult(false, true,
                "Die Serverantwort ist ausgeblieben. Die Buchung bleibt zur sicheren Prüfung vorgemerkt und wird nicht automatisch erneut gesendet.");
        }
        catch (Exception ex)
        {
            if (booking != null && !serverConfirmed)
                _tasks.FailTicketTimeBooking(booking);
            _logger.Error($"[ZnunyTimeBooking] ticketId={ticketId} taskId={task.Id} action=failed message={ex.Message}");
            if (serverConfirmed)
                return new TicketBookingResult(false, true, "Znuny hat die Buchung bestätigt, aber die lokale Bestätigung konnte nicht gespeichert werden. Beim nächsten Klick wird die Booking-ID sicher abgeglichen.");
            return new TicketBookingResult(false, false, $"Zeitbuchung fehlgeschlagen: {ex.Message}");
        }
    }

    public async Task<TicketBookingResult> CheckTicketTimeBookingAsync(TaskItem task, TicketTimeBooking booking)
    {
        try
        {
            var sessionId = await CreateSessionAsync();
            var ticket = await GetTicketAsync(booking.TicketId, sessionId, HashSessionId(sessionId));
            var articleId = ticket?.FindArticleIdContaining(BookingMarker(booking.BookingId));
            if (!string.IsNullOrWhiteSpace(articleId))
            {
                _tasks.CompleteTicketTimeBooking(booking, articleId);
                _logger.Info($"[ZnunyTimeBooking] ticketId={booking.TicketId} taskId={task.Id} bookingId={booking.BookingId} articleId={articleId} action=reconciled-manual");
                return new TicketBookingResult(true, false, "Buchung wurde in Znuny gefunden und lokal als gebucht bestätigt.");
            }

            _tasks.FailTicketTimeBooking(booking);
            return new TicketBookingResult(false, false, "Die Booking-ID wurde in den Ticketartikeln nicht gefunden. Eine erneute Buchung ist nur über den separaten Button möglich.");
        }
        catch (Exception ex)
        {
            _logger.Error($"[ZnunyTimeBooking] ticketId={booking.TicketId} taskId={task.Id} bookingId={booking.BookingId} action=reconciliation-failed message={ex.Message}");
            return new TicketBookingResult(false, true, $"Buchungsstatus konnte nicht geprüft werden: {ex.Message}");
        }
    }

    public async Task<TicketBookingResult> RetryTicketTimeBookingAsync(TaskItem task, TicketTimeBooking booking)
    {
        try
        {
            var sessionId = await CreateSessionAsync();
            var sessionHash = HashSessionId(sessionId);
            var ticket = await GetTicketAsync(booking.TicketId, sessionId, sessionHash);
            var existingArticleId = ticket?.FindArticleIdContaining(BookingMarker(booking.BookingId));
            if (!string.IsNullOrWhiteSpace(existingArticleId))
            {
                _tasks.CompleteTicketTimeBooking(booking, existingArticleId);
                return new TicketBookingResult(true, false, "Buchung war bereits vorhanden und wurde ohne erneutes TicketUpdate übernommen.");
            }

            _tasks.ResetTicketTimeBookingForRetry(booking);
            var timeUnit = booking.BookedMinutes;
            var payload = BuildTicketTimeBookingPayload(booking.TicketId, sessionId, booking, timeUnit);
            var route = ResolveTicketUpdateRoute(booking.TicketId);
            LogTicketUpdateRequest(route, booking.TicketId, booking, timeUnit);
            using var request = new HttpRequestMessage(HttpMethod.Post, Combine(_settings.Current.TicketSystemApiUrl, route))
            {
                Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
            };
            var response = await SendZnunyAsync(request, "TicketUpdateTimeBookingRetry", "[ZnunyTicketUpdateResponse]");
            EnsureTicketUpdateResponseIsInterpretable(response);
            var articleId = ExtractFirstValueRecursive(response.Body, "ArticleID");
            _tasks.CompleteTicketTimeBooking(booking, articleId);
            _logger.Info($"[ZnunyTimeBooking] ticketId={booking.TicketId} taskId={task.Id} bookingId={booking.BookingId} articleId={articleId} action=retried");
            return new TicketBookingResult(true, false, $"{booking.BookedMinutes:0.##} Min. erfolgreich erneut übertragen.");
        }
        catch (ZnunyApiException ex) when ((int)ex.StatusCode < 500)
        {
            _tasks.FailTicketTimeBooking(booking);
            return new TicketBookingResult(false, false, FormatApiError("Erneute Zeitbuchung fehlgeschlagen", ex));
        }
        catch (Exception ex)
        {
            _logger.Error($"[ZnunyTimeBooking] ticketId={booking.TicketId} taskId={task.Id} bookingId={booking.BookingId} action=retry-pending message={ex.Message}");
            return new TicketBookingResult(false, true, "Die erneute Übertragung ist unklar und bleibt zur Reconciliation auf Pending. Es erfolgt kein weiterer automatischer Versuch.");
        }
    }

    public async Task<(bool success, string message)> TestConnectionAsync()
    {
        LastError = string.Empty;

        try
        {
            var configError = ValidateConfiguration(requireAgentId: true);
            if (!string.IsNullOrWhiteSpace(configError))
                return (false, configError);

            var userId = GetConfiguredAgentId()!.Value;
            var sessionId = await CreateSessionAsync();
            var sessionHash = HashSessionId(sessionId);
            var (ownerIds, responsibleIds, uniqueIds) = await SearchAssignedTicketIdsAsync(userId, sessionId, sessionHash, includeOwner: true, includeResponsible: true);
            var duplicateCount = ownerIds.Count + responsibleIds.Count - uniqueIds.Count;

            var ticketGetStatus = "Nicht ausgeführt, keine Tickets gefunden.";
            if (uniqueIds.Count > 0)
            {
                var ticket = await GetTicketAsync(uniqueIds[0], sessionId, sessionHash);
                ticketGetStatus = ticket == null ? "Fehlgeschlagen, keine Ticketdaten in der Antwort." : "Erfolgreich";
            }

            return (true, $"Login/Authentifizierung: Erfolgreich\nTicketSearch-Route: {_settings.Current.TicketSystemTicketSearchMethod} {_settings.Current.TicketSystemTicketSearchRoute}\nAgenten-ID: {userId}\nOwner-Tickets: {ownerIds.Count}\nResponsible-Tickets: {responsibleIds.Count}\nEindeutige Tickets: {uniqueIds.Count}\nDoppelte Owner/Responsible-Treffer: {duplicateCount}\nTicketGet: {ticketGetStatus}\nMapping auf internen Task: Validiert");
        }
        catch (ZnunyApiException ex)
        {
            LogZnunyError(ex);
            return (false, FormatApiError("Znuny-Verbindungstest fehlgeschlagen", ex));
        }
        catch (Exception ex)
        {
            _logger.Error($"[ZnunyError] stage=Authentication errorCode={ex.HResult:X8} message={ex.Message}");
            return (false, $"Znuny-Verbindungstest fehlgeschlagen: {ex.Message}");
        }
    }

    public async Task<(bool success, string message)> TestRoutesAsync()
    {
        LastError = string.Empty;

        try
        {
            var configError = ValidateConfiguration(requireAgentId: true);
            if (!string.IsNullOrWhiteSpace(configError))
                return (false, configError);

            var userId = GetConfiguredAgentId()!.Value;
            var sessionId = await CreateSessionAsync();
            var sessionHash = HashSessionId(sessionId);
            try
            {
                var getIds = await SearchTicketsAsync("Owner", userId, "/Ticket", "GET", sessionId, sessionHash);
                _settings.Current.TicketSystemTicketSearchRoute = "/Ticket";
                _settings.Current.TicketSystemTicketSearchMethod = "GET";
                _settings.Current.TicketSystemTicketSearchAuthMode = "Session";
                _settings.Save();
                return (true, $"API-Routentest erfolgreich: GET /Ticket funktioniert. Owner-Tickets: {getIds.Count}. Route wurde gespeichert.");
            }
            catch (ZnunyApiException ex) when (IsRoutingError(ex))
            {
                LogZnunyError(ex);
            }

            var postIds = await SearchTicketsAsync("Owner", userId, "/Ticket/Search", "POST", sessionId, sessionHash);
            _settings.Current.TicketSystemTicketSearchRoute = "/Ticket/Search";
            _settings.Current.TicketSystemTicketSearchMethod = "POST";
            _settings.Current.TicketSystemTicketSearchAuthMode = "Session";
            _settings.Save();
            return (true, $"API-Routentest erfolgreich: POST /Ticket/Search funktioniert. Owner-Tickets: {postIds.Count}. Route wurde gespeichert.");
        }
        catch (ZnunyApiException ex)
        {
            LogZnunyError(ex);
            return (false, FormatApiError("API-Routentest fehlgeschlagen", ex));
        }
        catch (Exception ex)
        {
            _logger.Error($"[ZnunyError] stage=Authentication errorCode={ex.HResult:X8} message={ex.Message}");
            return (false, $"API-Routentest fehlgeschlagen: {ex.Message}");
        }
    }

    public async Task<(int created, int updated, int skipped)> SyncAssignedTicketsAsync(string reason)
    {
        if (!await _syncGate.WaitAsync(0))
        {
            _logger.Info($"[Znuny] Sync skipped because another run is active. reason={reason}");
            return (0, 0, 0);
        }

        LastError = string.Empty;
        var created = 0;
        var updated = 0;
        var skipped = 0;
        var hasTaskChanges = false;

        try
        {
            var configError = ValidateConfiguration(requireAgentId: true);
            if (!string.IsNullOrWhiteSpace(configError))
                return Fail3(configError);
            if (!_settings.Current.TicketSystemIncludeOwner && !_settings.Current.TicketSystemIncludeResponsible)
                return Fail3("Znuny Sync benötigt Owner oder Responsible als Suchkriterium.");

            var userId = GetConfiguredAgentId()!.Value;
            var sessionId = await CreateSessionAsync();
            var sessionHash = HashSessionId(sessionId);
            _logger.Info($"[Znuny] Sync start reason={reason} baseUrl='{SanitizeUrl(_settings.Current.TicketSystemApiUrl)}' auth={_settings.Current.TicketSystemTicketSearchAuthMode} onlyOpen={_settings.Current.TicketSystemOnlyOpenTickets} showClosed={_settings.Current.TicketSystemShowClosedTickets} includeOwner={_settings.Current.TicketSystemIncludeOwner} includeResponsible={_settings.Current.TicketSystemIncludeResponsible}");
            _logger.Info($"[ZnunyUser] source=ConfiguredSettings userId={userId}");

            var (ownerIds, responsibleIds, uniqueTicketIds) = await SearchAssignedTicketIdsAsync(
                userId,
                sessionId,
                sessionHash,
                _settings.Current.TicketSystemIncludeOwner,
                _settings.Current.TicketSystemIncludeResponsible);
            var existingGroups = _tasks.GetAllTasks()
                .Where(t => !string.IsNullOrWhiteSpace(ExtractZnunyTicketIdFromTask(t)))
                .GroupBy(ExtractZnunyTicketIdFromTask, StringComparer.OrdinalIgnoreCase)
                .ToList();
            var existing = existingGroups
                .Where(group => group.Count() == 1)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
            var ambiguousTicketIds = existingGroups
                .Where(group => group.Count() > 1)
                .Select(group => group.Key)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            foreach (var ambiguousTicketId in ambiguousTicketIds)
                _logger.Error($"[ZnunyTaskMapping] ticketId={ambiguousTicketId} action=skipped-ambiguous-task-mapping");

            // An open-only TicketSearch no longer returns a ticket as soon as it is closed.
            // Re-fetch already synchronized IDs explicitly so TicketGet can provide their
            // authoritative State/StateType and closure is never inferred from absence.
            var ticketIds = uniqueTicketIds
                .Concat(_settings.Current.TicketSystemOnlyOpenTickets ? existingGroups.Select(group => group.Key) : Array.Empty<string>())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            var duplicateCount = ownerIds.Count + responsibleIds.Count - uniqueTicketIds.Count;
            _logger.Info($"[ZnunySearchMerge] ownerTickets={ownerIds.Count} responsibleTickets={responsibleIds.Count} uniqueTickets={uniqueTicketIds.Count} existingTicketsRechecked={ticketIds.Count - uniqueTicketIds.Count} duplicateOwnerResponsibleTickets={duplicateCount}");

            foreach (var ticketId in ticketIds)
            {
                var ticket = await GetTicketAsync(ticketId, sessionId, sessionHash);
                if (ticket == null)
                {
                    skipped++;
                    continue;
                }

                if (ambiguousTicketIds.Contains(ticket.TicketID))
                {
                    skipped++;
                    continue;
                }

                if (existing.TryGetValue(ticket.TicketID, out var task))
                {
                    var metadataChanged = MapTicketToTask(ticket, task);
                    if (ticket.IsClosed)
                    {
                        if (task.Status == TaskStatus.Done)
                        {
                            if (metadataChanged)
                            {
                                _tasks.UpdateTask(task);
                                updated++;
                                hasTaskChanges = true;
                            }

                            LogAutoComplete(ticket, task, "already-completed");
                        }
                        else
                        {
                            // Use the same persistence path as the manual "Erledigt" action.
                            _tasks.MarkDone(task);
                            updated++;
                            hasTaskChanges = true;
                            LogAutoComplete(ticket, task, "completed");
                        }
                    }
                    else if (metadataChanged)
                    {
                        // Deliberately preserve Done: reopening in Znuny is not propagated back.
                        _tasks.UpdateTask(task);
                        updated++;
                        hasTaskChanges = true;
                        _logger.Info($"[ZnunyTaskUpdated] ticketId={ticket.TicketID} ticketNumber='{ticket.TicketNumber}' taskId={task.Id}");
                    }
                }
                else
                {
                    if (ticket.IsClosed && _settings.Current.TicketSystemOnlyOpenTickets && !_settings.Current.TicketSystemShowClosedTickets)
                    {
                        skipped++;
                        continue;
                    }

                    task = new TaskItem();
                    MapTicketToTask(ticket, task);
                    task.Status = ticket.IsClosed ? TaskStatus.Done : TaskStatus.Planned;
                    _tasks.CreateTask(task);
                    created++;
                    hasTaskChanges = true;
                    _logger.Info($"[ZnunyTaskCreated] ticketId={ticket.TicketID} ticketNumber='{ticket.TicketNumber}' taskId={task.Id}");
                }
            }

            _logger.Info($"[ZnunySyncFinished] created={created} updated={updated} skipped={skipped} totalTickets={ticketIds.Count}");
            return (created, updated, skipped);
        }
        catch (ZnunyApiException ex)
        {
            LastError = FormatApiError("Znuny Sync fehlgeschlagen", ex);
            LogZnunyError(ex);
            return (created, updated, skipped);
        }
        catch (Exception ex)
        {
            LastError = $"Znuny Sync fehlgeschlagen: {ex.Message}";
            _logger.Error($"[ZnunyError] stage=TaskSync errorCode={ex.HResult:X8} message={ex.Message}");
            return (created, updated, skipped);
        }
        finally
        {
            _syncGate.Release();
            if (hasTaskChanges)
                NotifyTasksChanged();
        }
    }

    private string ValidateConfiguration(bool requireAgentId)
    {
        if (string.IsNullOrWhiteSpace(_settings.Current.TicketSystemApiUrl))
            return "Znuny Server URL fehlt.";
        if (IsPlaceholderUrl(_settings.Current.TicketSystemApiUrl))
            return "Bitte die Znuny API-URL in den Einstellungen anpassen und SERVER durch den echten Hostnamen ersetzen.";
        if (string.IsNullOrWhiteSpace(_settings.Current.TicketSystemUsername))
            return "Znuny Benutzername fehlt.";
        if (string.IsNullOrWhiteSpace(_settings.GetTicketSystemPassword()))
            return "Znuny Passwort fehlt.";
        if (requireAgentId && _settings.Current.TicketSystemAgentId <= 0)
            return "Bitte eine Znuny Agenten-ID größer 0 in den Einstellungen eintragen.";

        return string.Empty;
    }

    private async Task<string> CreateSessionAsync()
    {
        var route = "/Session";
        var payload = new Dictionary<string, object?>
        {
            ["UserLogin"] = _settings.Current.TicketSystemUsername,
            ["Password"] = _settings.GetTicketSystemPassword()
        };
        using var request = new HttpRequestMessage(HttpMethod.Post, Combine(_settings.Current.TicketSystemApiUrl, route))
        {
            Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
        };

        _logger.Info($"[ZnunyLogin] method=POST route={route} payload={{UserLogin:'{_settings.Current.TicketSystemUsername}',Password:'***'}}");
        var result = await SendZnunyAsync(request, "SessionCreate", "[ZnunyLoginResponse]");
        using var doc = JsonDocument.Parse(result.Body);
        var sessionId = FirstString(doc.RootElement, "SessionID");
        if (string.IsNullOrWhiteSpace(sessionId))
            throw new ZnunyApiException("SessionCreate", result.StatusCode, "Protocol", "SessionCreate response contains no SessionID.", result.Body);

        _logger.Info($"[ZnunySession] sessionCreated=True sessionHash={HashSessionId(sessionId)}");
        return sessionId;
    }

    private async Task<JsonDocument> GetSessionAsync(string sessionId, string sessionHash)
    {
        var route = $"/Session/SessionID={Uri.EscapeDataString(sessionId)}";
        using var request = new HttpRequestMessage(HttpMethod.Get, Combine(_settings.Current.TicketSystemApiUrl, route));
        _logger.Info($"[ZnunySession] method=GET route=/Session/SessionID=*** sessionHash={sessionHash} diagnostic=True");
        var result = await SendZnunyAsync(request, "SessionGetDiagnostic", "[ZnunySessionResponse]");
        var doc = JsonDocument.Parse(result.Body);
        LogSessionKeys(doc.RootElement);
        return doc;
    }

    private static int? ResolveUserId(JsonDocument sessionData)
        => FindSessionValue(sessionData.RootElement, "UserID", "UserId") ?? FindInteger(sessionData.RootElement, "UserID", "UserId");

    private async Task<int?> TryResolveUserIdFromSessionAsync(string sessionId, string sessionHash)
    {
        try
        {
            using var sessionData = await GetSessionAsync(sessionId, sessionHash);
            return ResolveUserId(sessionData);
        }
        catch (ZnunyApiException ex)
        {
            LogZnunyError(ex);
            return null;
        }
        catch (Exception ex)
        {
            _logger.Error($"[ZnunyError] stage=SessionGet errorCode={ex.HResult:X8} message={ex.Message}");
            return null;
        }
    }

    private int? GetConfiguredAgentId()
        => _settings.Current.TicketSystemAgentId > 0 ? _settings.Current.TicketSystemAgentId : null;

    private async Task<(List<string> ownerIds, List<string> responsibleIds, List<string> uniqueIds)> SearchAssignedTicketIdsAsync(
        int userId,
        string sessionId,
        string sessionHash,
        bool includeOwner,
        bool includeResponsible)
    {
        var ownerIds = includeOwner
            ? await SearchRoleTicketIdsWithOpenCompatibilityAsync("Owner", userId, sessionId, sessionHash)
            : new List<string>();
        var responsibleIds = includeResponsible
            ? await SearchRoleTicketIdsWithOpenCompatibilityAsync("Responsible", userId, sessionId, sessionHash)
            : new List<string>();
        var uniqueIds = ownerIds
            .Concat(responsibleIds)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        return (ownerIds, responsibleIds, uniqueIds);
    }

    private async Task<List<string>> SearchRoleTicketIdsWithOpenCompatibilityAsync(string role, int userId, string sessionId, string sessionHash)
    {
        var onlyOpen = _settings.Current.TicketSystemOnlyOpenTickets && !_settings.Current.TicketSystemShowClosedTickets;
        var filteredIds = await SearchTicketsAsync(role, userId, _settings.Current.TicketSystemTicketSearchRoute, _settings.Current.TicketSystemTicketSearchMethod, sessionId, sessionHash, onlyOpen);

        if (!onlyOpen)
            return filteredIds;

        var unfilteredIds = await SearchTicketsAsync(role, userId, _settings.Current.TicketSystemTicketSearchRoute, _settings.Current.TicketSystemTicketSearchMethod, sessionId, sessionHash, onlyOpenOverride: false);
        var mergedIds = filteredIds
            .Concat(unfilteredIds)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        _logger.Info($"[ZnunySearchCompatibility] role={role} stateTypeOpenTickets={filteredIds.Count} unfilteredTickets={unfilteredIds.Count} mergedTickets={mergedIds.Count}");
        return mergedIds;
    }

    private async Task<List<string>> SearchTicketsAsync(string role, int userId, string route, string method, string sessionId, string sessionHash, bool? onlyOpenOverride = null)
    {
        route = NormalizeRouteValue(route, "/Ticket");
        method = string.Equals(method, "GET", StringComparison.OrdinalIgnoreCase) ? "GET" : "POST";
        var isOwner = role == "Owner";
        var stage = isOwner ? "TicketSearchOwner" : "TicketSearchResponsible";
        var logTag = isOwner ? "[ZnunySearchOwner]" : "[ZnunySearchResponsible]";
        var onlyOpen = onlyOpenOverride ?? (_settings.Current.TicketSystemOnlyOpenTickets && !_settings.Current.TicketSystemShowClosedTickets);
        if (method == "GET" && string.Equals(_settings.Current.TicketSystemTicketSearchAuthMode, "Direct", StringComparison.OrdinalIgnoreCase))
            throw new ZnunyApiException(stage, HttpStatusCode.BadRequest, "Configuration", "Direkte Authentifizierung für TicketSearch darf nicht mit GET verwendet werden, weil Credentials in URLs/Logs landen würden.", string.Empty);

        var payload = BuildSearchPayload(isOwner, userId, onlyOpen, sessionId);
        using var request = BuildSearchRequest(method, route, payload);

        _logger.Info($"{logTag} method={method} route={route} userId={userId} onlyOpen={onlyOpen} sessionHash={sessionHash} payload={FormatSearchPayloadForLog(isOwner, userId, onlyOpen, _settings.Current.TicketSystemTicketSearchAuthMode, _settings.Current.TicketSystemUsername)}");
        var result = await SendZnunyAsync(request, stage, isOwner ? "[ZnunySearchOwnerResponse]" : "[ZnunySearchResponsibleResponse]");
        var ticketIds = ExtractTicketIdsStrict(result.Body, stage).ToList();
        _logger.Info($"{logTag} method={method} route={route} userId={userId} onlyOpen={onlyOpen} status={(int)result.StatusCode} ticketCount={ticketIds.Count}");
        return ticketIds;
    }

    private async Task<ZnunyTicket?> GetTicketAsync(string ticketId, string sessionId, string sessionHash)
    {
        var route = NormalizeRouteValue(_settings.Current.TicketSystemTicketGetRouteTemplate, "/Ticket/{TicketID}")
            .Replace("{TicketID}", Uri.EscapeDataString(ticketId), StringComparison.OrdinalIgnoreCase);
        var query = new Dictionary<string, string>();
        if (string.Equals(_settings.Current.TicketSystemTicketGetAuthMode, "Direct", StringComparison.OrdinalIgnoreCase))
        {
            throw new ZnunyApiException("TicketGet", HttpStatusCode.BadRequest, "Configuration", "Direkte Authentifizierung für TicketGet per GET ist deaktiviert, damit das Passwort nicht in URL-, Proxy- oder Server-Logs gelangt. Bitte Session verwenden.", string.Empty);
        }
        else
        {
            query["SessionID"] = sessionId;
        }
        query["AllArticles"] = "1";
        query["DynamicFields"] = "1";
        var url = Combine(_settings.Current.TicketSystemApiUrl, route) + ToQueryString(query);
        using var request = new HttpRequestMessage(HttpMethod.Get, url);

        _logger.Info($"[ZnunyTicket] method=GET route={route} ticketId={ticketId} auth={_settings.Current.TicketSystemTicketGetAuthMode} requestedDynamicFields=true allArticles=true sessionHash={sessionHash}");
        var result = await SendZnunyAsync(request, "TicketGet", "[ZnunyTicketResponse]");
        using var doc = JsonDocument.Parse(result.Body);
        ThrowIfApiError(doc.RootElement, "TicketGet");
        var ticketElement = FindFirstTicketElement(doc.RootElement);
        if (!ticketElement.HasValue)
            throw new ZnunyApiException("TicketGet", result.StatusCode, "Protocol", "TicketGet response contains no Ticket object.", result.Body);

        var ticket = ZnunyTicket.FromJson(ticketElement.Value, _settings.Current.TicketSystemWebUrl, doc.RootElement);
        _logger.Info($"[ZnunyFirstArticle] ticketId={ticket.TicketID} articleCount={ticket.ArticleCount} selectedArticleId='{ticket.FirstArticleId}' senderType='{ticket.FirstArticleSenderType}' created='{ticket.FirstArticleCreated}' bodyLength={ticket.FirstArticleBody.Length}");
        return ticket;
    }

    private async Task<ZnunyHttpResult> SendZnunyAsync(HttpRequestMessage request, string stage, string responseLogTag)
    {
        using var response = await _client.SendAsync(request);
        var body = await response.Content.ReadAsStringAsync();
        var contentType = response.Content.Headers.ContentType?.ToString() ?? string.Empty;
        _logger.Info($"{responseLogTag} status={(int)response.StatusCode} contentType='{contentType}' body={Truncate(RedactSecrets(body))}");

        if (!response.IsSuccessStatusCode)
        {
            var (errorCode, errorMessage) = ExtractApiError(body);
            throw new ZnunyApiException(stage, response.StatusCode, errorCode, string.IsNullOrWhiteSpace(errorMessage) ? response.ReasonPhrase ?? "HTTP error" : errorMessage, body);
        }

        if (TryParseJson(body, out var doc))
        {
            using (doc)
            {
                ThrowIfApiError(doc.RootElement, stage, response.StatusCode, body);
            }
        }

        return new ZnunyHttpResult(response.StatusCode, contentType, body);
    }

    private HttpRequestMessage BuildSearchRequest(string method, string route, Dictionary<string, object?> payload)
    {
        if (method == "GET")
        {
            var query = payload.ToDictionary(kvp => kvp.Key, kvp => FormatQueryValue(kvp.Value));
            return new HttpRequestMessage(HttpMethod.Get, Combine(_settings.Current.TicketSystemApiUrl, route) + ToQueryString(query));
        }

        var json = JsonSerializer.Serialize(payload);
        return new HttpRequestMessage(HttpMethod.Post, Combine(_settings.Current.TicketSystemApiUrl, route))
        {
            Content = new StringContent(json, Encoding.UTF8, "application/json")
        };
    }

    private Dictionary<string, object?> BuildSearchPayload(bool owner, int userId, bool onlyOpen, string sessionId)
    {
        var payload = new Dictionary<string, object?>
        {
            [owner ? "OwnerIDs" : "ResponsibleIDs"] = userId
        };

        if (string.Equals(_settings.Current.TicketSystemTicketSearchAuthMode, "Direct", StringComparison.OrdinalIgnoreCase))
        {
            payload["UserLogin"] = _settings.Current.TicketSystemUsername;
            payload["Password"] = _settings.GetTicketSystemPassword();
        }
        else
        {
            payload["SessionID"] = sessionId;
        }

        if (onlyOpen)
            payload["StateType"] = "Open";

        return payload;
    }

    private static bool MapTicketToTask(ZnunyTicket ticket, TaskItem task)
    {
        var title = $"[{ticket.TicketNumber}] {ticket.Title}".Trim();
        var description = ticket.ToDescription();
        var tags = $"Znuny;ZnunyTicketID:{ticket.TicketID};ZnunyTicketNumber:{ticket.TicketNumber}";
        var changed = !string.Equals(task.Title, title, StringComparison.Ordinal)
                      || !string.Equals(task.Description, description, StringComparison.Ordinal)
                      || !string.Equals(task.TicketUrl, ticket.WebUrl, StringComparison.Ordinal)
                      || !string.Equals(task.Tags, tags, StringComparison.Ordinal);

        task.Title = title;
        task.Description = description;
        task.TicketUrl = ticket.WebUrl;
        task.Tags = tags;
        return changed;
    }

    private void LogAutoComplete(ZnunyTicket ticket, TaskItem task, string action)
        => _logger.Info($"[ZnunyAutoComplete] ticketId={ticket.TicketID} ticketNumber='{ticket.TicketNumber}' taskId={task.Id} ticketState='{ticket.State}' ticketStateType='{ticket.StateType}' action={action}");

    private void NotifyTasksChanged()
    {
        try
        {
            TasksChanged?.Invoke();
        }
        catch (Exception ex)
        {
            _logger.Error($"[ZnunyTaskRefresh] action=failed errorCode={ex.HResult:X8} message={ex.Message}");
        }
    }

    private static string ExtractZnunyTicketIdFromTask(TaskItem task)
    {
        var parts = (task.Tags ?? string.Empty).Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var id = parts.FirstOrDefault(p => p.StartsWith("ZnunyTicketID:", StringComparison.OrdinalIgnoreCase));
        return id?.Split(':', 2).ElementAtOrDefault(1) ?? string.Empty;
    }

    private static string ExtractZnunyTicketNumberFromTask(TaskItem task)
    {
        var parts = (task.Tags ?? string.Empty).Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var number = parts.FirstOrDefault(p => p.StartsWith("ZnunyTicketNumber:", StringComparison.OrdinalIgnoreCase));
        return number?.Split(':', 2).ElementAtOrDefault(1) ?? string.Empty;
    }

    private Dictionary<string, object?> BuildTicketTimeBookingPayload(
        string ticketId,
        string sessionId,
        TicketTimeBooking booking,
        decimal timeUnit)
    {
        var body = $"{booking.ShortDescription}\n\n{BookingMarker(booking.BookingId)}";
        var payload = new Dictionary<string, object?>
        {
            ["SessionID"] = sessionId,
            ["TicketID"] = ticketId,
            ["Article"] = new Dictionary<string, object?>
            {
                ["Subject"] = "TaskTool Zeitbuchung",
                ["Body"] = body,
                ["ContentType"] = "text/plain; charset=utf-8",
                ["MimeType"] = "text/plain",
                ["Charset"] = "utf-8",
                ["SenderType"] = "agent",
                ["CommunicationChannel"] = "Internal",
                ["IsVisibleForCustomer"] = 0,
                ["TimeUnit"] = timeUnit
            }
        };

        var dynamicFields = new List<Dictionary<string, string>>();
        AddDynamicField(dynamicFields, _settings.Current.TicketSystemCostCenterFieldName, booking.CostCenter);
        AddDynamicField(dynamicFields, _settings.Current.TicketSystemOrderFieldName, booking.Order);
        if (dynamicFields.Count > 0)
            payload["DynamicField"] = dynamicFields;
        return payload;
    }

    private void LogTicketUpdateRequest(string route, string ticketId, TicketTimeBooking booking, decimal timeUnit)
    {
        var dynamicFields = new List<string>();
        if (!string.IsNullOrWhiteSpace(_settings.Current.TicketSystemCostCenterFieldName) && !string.IsNullOrWhiteSpace(booking.CostCenter))
            dynamicFields.Add($"{LogValue(_settings.Current.TicketSystemCostCenterFieldName)}={LogValue(booking.CostCenter)}");
        if (!string.IsNullOrWhiteSpace(_settings.Current.TicketSystemOrderFieldName) && !string.IsNullOrWhiteSpace(booking.Order))
            dynamicFields.Add($"{LogValue(_settings.Current.TicketSystemOrderFieldName)}={LogValue(booking.Order)}");

        _logger.Info($"[ZnunyTicketUpdateRequest] route={route} ticketId={ticketId} articleSubject='TaskTool Zeitbuchung' senderType='agent' channel='Internal' visibleForCustomer=0 timeUnit={timeUnit:0.####} dynamicFields=[{string.Join(',', dynamicFields)}]");
    }

    private string ResolveTicketUpdateRoute(string ticketId)
    {
        var template = NormalizeRouteValue(
            _settings.Current.TicketSystemTicketUpdateRoute,
            AppSettings.DefaultTicketSystemTicketUpdateRoute);
        var route = template.Contains("{TicketID}", StringComparison.OrdinalIgnoreCase)
            ? template.Replace("{TicketID}", Uri.EscapeDataString(ticketId), StringComparison.OrdinalIgnoreCase)
            : template;
        if (route.Contains("{TicketID}", StringComparison.OrdinalIgnoreCase))
        {
            _logger.Error($"[ZnunyTicketUpdateRequest] ticketId={ticketId} action=blocked reason=UnresolvedTicketIDPlaceholder");
            throw new InvalidOperationException("Der {TicketID}-Platzhalter der TicketUpdate-Route konnte nicht aufgelöst werden.");
        }
        return route;
    }

    private static string LogValue(string value)
        => value.Replace("\r", string.Empty, StringComparison.Ordinal).Replace("\n", string.Empty, StringComparison.Ordinal).Replace("'", string.Empty, StringComparison.Ordinal);

    private static void EnsureTicketUpdateResponseIsInterpretable(ZnunyHttpResult response)
    {
        using var doc = JsonDocument.Parse(response.Body);
        ThrowIfApiError(doc.RootElement, "TicketUpdate", response.StatusCode, response.Body);
    }

    private static void AddDynamicField(List<Dictionary<string, string>> fields, string name, string value)
    {
        if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(value))
            return;
        fields.Add(new Dictionary<string, string> { ["Name"] = name.Trim(), ["Value"] = value.Trim() });
    }

    private static string BookingMarker(string bookingId) => $"TaskTool-Booking-ID: {bookingId}";

    private static IReadOnlyList<TicketFieldOption> ParseConfiguredOptions(string configured)
    {
        return (configured ?? string.Empty)
            .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(entry => entry.Split('=', 2, StringSplitOptions.TrimEntries))
            .Where(parts => !string.IsNullOrWhiteSpace(parts[0]))
            .Select(parts => new TicketFieldOption(parts[0], parts.Length > 1 && !string.IsNullOrWhiteSpace(parts[1]) ? parts[1] : parts[0]))
            .GroupBy(option => option.Key, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .ToList();
    }

    private async Task<IReadOnlyDictionary<string, IReadOnlyList<TicketFieldOption>>> GetDynamicFieldOptionsAsync(string sessionId, bool forceRefresh)
    {
        if (!forceRefresh && _dynamicFieldOptionsCacheValid && DateTime.UtcNow < _dynamicFieldOptionsCacheExpiresUtc)
            return _dynamicFieldOptionsCache;

        await _dynamicFieldOptionsGate.WaitAsync();
        try
        {
            if (!forceRefresh && _dynamicFieldOptionsCacheValid && DateTime.UtcNow < _dynamicFieldOptionsCacheExpiresUtc)
                return _dynamicFieldOptionsCache;

            var names = new[] { _settings.Current.TicketSystemCostCenterFieldName, _settings.Current.TicketSystemOrderFieldName }
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            var parsed = new Dictionary<string, IReadOnlyList<TicketFieldOption>>(StringComparer.OrdinalIgnoreCase);
            foreach (var name in names)
            {
                try
                {
                    parsed[name] = await LoadDynamicFieldOptionsAsync(sessionId, name);
                }
                catch (Exception ex)
                {
                    parsed[name] = Array.Empty<TicketFieldOption>();
                    _logger.Error($"[ZnunyDynamicFieldOptions] field='{name}' optionCount=0 source=ConfiguredFallback message={ex.Message}");
                }
            }
            _dynamicFieldOptionsCache = parsed;
            _dynamicFieldOptionsCacheExpiresUtc = DateTime.UtcNow.AddMinutes(30);
            _dynamicFieldOptionsCacheValid = true;
            return _dynamicFieldOptionsCache;
        }
        catch (Exception ex)
        {
            _logger.Error($"[ZnunyDynamicFieldOptions] source=ConfiguredFallback message={ex.Message}");
            _dynamicFieldOptionsCache = new Dictionary<string, IReadOnlyList<TicketFieldOption>>(StringComparer.OrdinalIgnoreCase);
            _dynamicFieldOptionsCacheExpiresUtc = DateTime.UtcNow.AddMinutes(5);
            _dynamicFieldOptionsCacheValid = true;
            return _dynamicFieldOptionsCache;
        }
        finally
        {
            _dynamicFieldOptionsGate.Release();
        }
    }

    private async Task<IReadOnlyList<TicketFieldOption>> LoadDynamicFieldOptionsAsync(string sessionId, string fieldName)
    {
        var template = NormalizeRouteValue(
            _settings.Current.TicketSystemDynamicFieldOptionsRoute,
            "/Ticket/DynamicField/{FieldName}/Options");
        if (!template.Contains("{FieldName}", StringComparison.OrdinalIgnoreCase))
        {
            _logger.Error($"[ZnunyDynamicFieldOptionsRequest] field='{fieldName}' action=blocked reason=MissingFieldNamePlaceholder");
            throw new InvalidOperationException("Die DynamicField-Options-Route enthält keinen {FieldName}-Platzhalter.");
        }

        var route = template.Replace(
            "{FieldName}",
            Uri.EscapeDataString(fieldName),
            StringComparison.OrdinalIgnoreCase);
        if (route.Contains("{FieldName}", StringComparison.OrdinalIgnoreCase))
        {
            _logger.Error($"[ZnunyDynamicFieldOptionsRequest] field='{fieldName}' action=blocked reason=UnresolvedFieldNamePlaceholder");
            throw new InvalidOperationException("Der {FieldName}-Platzhalter konnte nicht aufgelöst werden.");
        }

        _logger.Info($"[ZnunyDynamicFieldOptionsRequest] field='{fieldName}' route='{route}'");
        var query = new Dictionary<string, string> { ["SessionID"] = sessionId };
        using var request = new HttpRequestMessage(HttpMethod.Get, Combine(_settings.Current.TicketSystemApiUrl, route) + ToQueryString(query));
        var response = await SendZnunyAsync(request, "DynamicFieldOptions", "[ZnunyDynamicFieldOptionsResponse]");
        var options = ParseDynamicFieldOptionsResponse(response.Body, fieldName);
        _logger.Info($"[ZnunyDynamicFieldOptions] field='{fieldName}' optionCount={options.Count} source=Znuny");
        return options;
    }

    private IReadOnlyList<TicketFieldOption> GetFieldOptions(
        IReadOnlyDictionary<string, IReadOnlyList<TicketFieldOption>> fields,
        string fieldName,
        string configuredFallback)
    {
        if (fields.TryGetValue(fieldName, out var options) && options.Count > 0)
            return options;
        var fallback = ParseConfiguredOptions(configuredFallback);
        _logger.Info($"[ZnunyDynamicFieldOptions] field='{fieldName}' optionCount={fallback.Count} source=ConfiguredFallback");
        return fallback;
    }

    private void LogDynamicFieldSelection(string ticketId, string fieldName, string currentKey, IReadOnlyList<TicketFieldOption> options)
    {
        var selectedKey = string.IsNullOrWhiteSpace(currentKey) && options.Any(option => option.Key == "00000") ? "00000" : currentKey;
        var display = options.FirstOrDefault(option => string.Equals(option.Key, selectedKey, StringComparison.OrdinalIgnoreCase))?.DisplayText
                      ?? selectedKey;
        _logger.Info($"[ZnunyDynamicFieldSelection] ticketId={ticketId} field='{fieldName}' key='{selectedKey}' displayValue='{display}'");
    }

    private static IReadOnlyList<TicketFieldOption> ParseDynamicFieldOptionsResponse(string json, string requestedFieldName)
    {
        using var doc = JsonDocument.Parse(json);
        if (!TryGetPropertyCaseInsensitive(doc.RootElement, "Field", out var field))
        {
            if (!TryGetPropertyCaseInsensitive(doc.RootElement, "Data", out var data)
                || !TryGetPropertyCaseInsensitive(data, "Field", out field))
                return Array.Empty<TicketFieldOption>();
        }
        var name = FirstString(field, "Name");
        if (!string.IsNullOrWhiteSpace(name) && !string.Equals(name, requestedFieldName, StringComparison.OrdinalIgnoreCase))
            return Array.Empty<TicketFieldOption>();
        if (!TryGetPropertyCaseInsensitive(field, "Options", out var values) || values.ValueKind != JsonValueKind.Array)
            return Array.Empty<TicketFieldOption>();
        return values.EnumerateArray()
            .Select(option => new TicketFieldOption(FirstString(option, "Key"), FirstString(option, "Value")))
            .Where(option => !string.IsNullOrWhiteSpace(option.Key))
            .GroupBy(option => option.Key, StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .ToList();
    }

    private static string ExtractFirstValueRecursive(string json, string propertyName)
    {
        if (!TryParseJson(json, out var doc)) return string.Empty;
        using (doc)
        {
            return FindFirstValueRecursive(doc.RootElement, propertyName);
        }
    }

    private static string FindFirstValueRecursive(JsonElement element, string propertyName)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            {
                if (string.Equals(property.Name, propertyName, StringComparison.OrdinalIgnoreCase))
                    return TicketIdToString(property.Value);
                var nested = FindFirstValueRecursive(property.Value, propertyName);
                if (!string.IsNullOrWhiteSpace(nested)) return nested;
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in element.EnumerateArray())
            {
                var nested = FindFirstValueRecursive(item, propertyName);
                if (!string.IsNullOrWhiteSpace(nested)) return nested;
            }
        }
        return string.Empty;
    }

    private void LogSessionKeys(JsonElement root)
    {
        var keys = CollectSessionKeys(root).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(k => k).ToList();
        var knownValues = new[]
        {
            "UserID", "UserId", "UserLogin", "UserEmail", "UserFirstname", "UserLastname"
        }.Select(key => $"{key}='{FindSessionString(root, key)}'");

        _logger.Info($"[ZnunySession] SessionGet keys=[{string.Join(",", keys)}] knownFields={{ {string.Join(", ", knownValues)} }}");
    }

    private static IEnumerable<string> CollectSessionKeys(JsonElement root)
    {
        if (!root.TryGetProperty("SessionData", out var data))
            yield break;

        if (data.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in data.EnumerateObject())
                yield return property.Name;
            yield break;
        }

        if (data.ValueKind != JsonValueKind.Array)
            yield break;

        foreach (var item in data.EnumerateArray())
        {
            var key = FirstString(item, "Key");
            if (!string.IsNullOrWhiteSpace(key))
                yield return key;
        }
    }

    private static string FindSessionString(JsonElement root, string key)
    {
        if (!root.TryGetProperty("SessionData", out var data))
            return string.Empty;

        if (data.ValueKind == JsonValueKind.Object)
            return FirstString(data, key);

        if (data.ValueKind != JsonValueKind.Array)
            return string.Empty;

        foreach (var item in data.EnumerateArray())
        {
            if (string.Equals(FirstString(item, "Key"), key, StringComparison.OrdinalIgnoreCase))
                return FirstString(item, "Value");
        }

        return string.Empty;
    }

    private static bool ContainsError(JsonElement root, out string errorCode, out string errorMessage)
    {
        errorCode = string.Empty;
        errorMessage = string.Empty;

        if (!TryGetPropertyCaseInsensitive(root, "Error", out var error) || error.ValueKind != JsonValueKind.Object)
            return false;

        errorCode = FirstString(error, "ErrorCode");
        errorMessage = FirstString(error, "ErrorMessage");
        return true;
    }

    private static string HashSessionId(string sessionId)
    {
        var bytes = SHA256.HashData(Encoding.UTF8.GetBytes(sessionId));
        return Convert.ToHexString(bytes)[..12];
    }

    private static void ThrowIfApiError(JsonElement root, string stage, HttpStatusCode statusCode = HttpStatusCode.OK, string responseBody = "")
    {
        if (ContainsError(root, out var errorCode, out var errorMessage))
            throw new ZnunyApiException(stage, statusCode, errorCode, errorMessage, responseBody);
    }

    private static IEnumerable<string> ExtractTicketIdsStrict(string responseBody, string stage)
    {
        if (!TryParseJson(responseBody, out var doc))
            throw new ZnunyApiException(stage, HttpStatusCode.OK, "Protocol", "TicketSearch response is not valid JSON.", responseBody);

        using (doc)
        {
            ThrowIfApiError(doc.RootElement, stage, HttpStatusCode.OK, responseBody);
            if (TryGetPropertyCaseInsensitive(doc.RootElement, "TicketIDs", out var ids) || TryGetPropertyCaseInsensitive(doc.RootElement, "TicketID", out ids))
                return ExtractTicketIdValues(ids).ToList();
        }

        throw new ZnunyApiException(stage, HttpStatusCode.OK, "Protocol", "TicketSearch response contains neither TicketID nor TicketIDs.", responseBody);
    }

    private static IEnumerable<string> ExtractTicketIdValues(JsonElement value)
    {
        if (value.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in value.EnumerateArray())
            {
                var id = TicketIdToString(item);
                if (!string.IsNullOrWhiteSpace(id))
                    yield return id;
            }

            yield break;
        }

        var single = TicketIdToString(value);
        if (!string.IsNullOrWhiteSpace(single))
            yield return single;
    }

    private static string TicketIdToString(JsonElement value)
        => value.ValueKind switch
        {
            JsonValueKind.String => value.GetString() ?? string.Empty,
            JsonValueKind.Number => value.ToString(),
            _ => string.Empty
        };

    private static (string errorCode, string errorMessage) ExtractApiError(string responseBody)
    {
        if (!TryParseJson(responseBody, out var doc))
            return ("HTTP", responseBody);

        using (doc)
        {
            return ContainsError(doc.RootElement, out var errorCode, out var errorMessage)
                ? (errorCode, errorMessage)
                : ("HTTP", responseBody);
        }
    }

    private static bool TryParseJson(string json, out JsonDocument doc)
    {
        try
        {
            doc = JsonDocument.Parse(json);
            return true;
        }
        catch
        {
            doc = null!;
            return false;
        }
    }

    private static bool TryGetPropertyCaseInsensitive(JsonElement root, string name, out JsonElement value)
    {
        if (root.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in root.EnumerateObject())
            {
                if (string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    value = property.Value;
                    return true;
                }
            }
        }

        value = default;
        return false;
    }

    private static string NormalizeRouteValue(string? route, string defaultRoute)
    {
        if (string.IsNullOrWhiteSpace(route))
            return defaultRoute;

        route = route.Trim();
        return route.StartsWith('/') ? route : "/" + route;
    }

    private static string FormatQueryValue(object? value)
        => value switch
        {
            null => string.Empty,
            string text => text,
            int number => number.ToString(),
            int[] values => string.Join(",", values),
            IEnumerable<int> values => string.Join(",", values),
            _ => value.ToString() ?? string.Empty
        };

    private static string ToQueryString(Dictionary<string, string> query)
        => "?" + string.Join("&", query.Select(kvp => $"{Uri.EscapeDataString(kvp.Key)}={Uri.EscapeDataString(kvp.Value)}"));

    private static string FormatSearchPayloadForLog(bool owner, int userId, bool onlyOpen, string authMode, string userLogin)
    {
        var idName = owner ? "OwnerIDs" : "ResponsibleIDs";
        var state = onlyOpen ? ",StateType:'Open'" : string.Empty;
        var auth = string.Equals(authMode, "Direct", StringComparison.OrdinalIgnoreCase)
            ? $"UserLogin:'{userLogin}',Password:'***'"
            : "SessionID:'***'";
        return $"{{{auth},{idName}:{userId}{state}}}";
    }

    private static bool IsRoutingError(ZnunyApiException ex)
        => ex.StatusCode is HttpStatusCode.NotFound or HttpStatusCode.MethodNotAllowed
           || ex.ErrorMessage.Contains("operation not found", StringComparison.OrdinalIgnoreCase)
           || ex.ErrorMessage.Contains("could not determine operation", StringComparison.OrdinalIgnoreCase)
           || ex.ErrorMessage.Contains("no route", StringComparison.OrdinalIgnoreCase);

    private void LogZnunyError(ZnunyApiException ex)
        => _logger.Error($"[ZnunyError] stage={ex.Stage} httpStatus={(int)ex.StatusCode} errorCode={ex.ErrorCode} message={ex.ErrorMessage} response={Truncate(RedactSecrets(ex.ResponseBody))}");

    private string FormatApiError(string title, ZnunyApiException ex)
        => $"{title}\nStufe: {ex.Stage}\nHTTP-Status: {(int)ex.StatusCode}\nZnuny ErrorCode: {ex.ErrorCode}\nZnuny ErrorMessage: {ex.ErrorMessage}\nResponse: {Truncate(RedactSecrets(ex.ResponseBody))}";

    private sealed record ZnunyHttpResult(HttpStatusCode StatusCode, string ContentType, string Body);

    private static IEnumerable<string> ExtractTicketIds(JsonElement root)
    {
        if (root.TryGetProperty("TicketIDs", out var ids) && ids.ValueKind == JsonValueKind.Array)
            return ids.EnumerateArray().Select(v => v.ToString()).Where(v => !string.IsNullOrWhiteSpace(v)).ToList();
        if (root.TryGetProperty("TicketID", out var id) && id.ValueKind == JsonValueKind.Array)
            return id.EnumerateArray().Select(v => v.ToString()).Where(v => !string.IsNullOrWhiteSpace(v)).ToList();
        if (root.TryGetProperty("TicketID", out var singleId) && singleId.ValueKind is JsonValueKind.String or JsonValueKind.Number)
            return new[] { singleId.ToString() };
        return Array.Empty<string>();
    }

    private static JsonElement? FindFirstTicketElement(JsonElement root)
    {
        if (root.TryGetProperty("Ticket", out var tickets))
        {
            if (tickets.ValueKind == JsonValueKind.Array && tickets.GetArrayLength() > 0) return tickets[0];
            if (tickets.ValueKind == JsonValueKind.Object) return tickets;
        }

        return null;
    }

    private static int? FindSessionValue(JsonElement root, params string[] keys)
    {
        if (!root.TryGetProperty("SessionData", out var data))
            return null;

        if (data.ValueKind == JsonValueKind.Object)
            return FindInteger(data, keys);

        if (data.ValueKind != JsonValueKind.Array)
            return null;

        foreach (var item in data.EnumerateArray())
        {
            var key = FirstString(item, "Key");
            if (!keys.Contains(key, StringComparer.OrdinalIgnoreCase)) continue;
            if (int.TryParse(FirstString(item, "Value"), out var value)) return value;
        }

        return null;
    }

    private static int? FindInteger(JsonElement item, params string[] names)
    {
        if (item.ValueKind != JsonValueKind.Object)
            return null;

        foreach (var name in names)
        {
            foreach (var property in item.EnumerateObject())
            {
                if (!string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase)) continue;
                if (property.Value.ValueKind == JsonValueKind.Number && property.Value.TryGetInt32(out var number)) return number;
                if (int.TryParse(FirstString(item, name), out var parsed)) return parsed;
            }
        }

        return null;
    }

    private static string FirstString(JsonElement item, params string[] names)
    {
        if (item.ValueKind != JsonValueKind.Object) return string.Empty;
        foreach (var name in names)
        {
            foreach (var property in item.EnumerateObject())
            {
                if (!string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase)) continue;

                var value = property.Value;
                if (value.ValueKind == JsonValueKind.String) return value.GetString() ?? string.Empty;
                if (value.ValueKind is JsonValueKind.Number or JsonValueKind.True or JsonValueKind.False) return value.ToString();
            }
        }
        return string.Empty;
    }

    private static bool IsPlaceholderUrl(string value)
        => value.Contains("SERVER", StringComparison.OrdinalIgnoreCase);

    private static string Combine(string baseUrl, string relative) => $"{baseUrl.TrimEnd('/')}/{relative.TrimStart('/')}";
    private static string SanitizeUrl(string value)
    {
        var sanitized = value.Replace("Password=", "Password=***", StringComparison.OrdinalIgnoreCase);
        return Regex.Replace(sanitized, "SessionID=[^&/\\s]+", "SessionID=***", RegexOptions.IgnoreCase);
    }
    private static string Truncate(string value) => value.Length <= 3000 ? value : value[..3000] + "...";
    private string RedactSecrets(string value)
    {
        var redacted = Regex.Replace(value, "\"SessionID\"\\s*:\\s*\"[^\"]+\"", "\"SessionID\":\"***\"", RegexOptions.IgnoreCase);
        redacted = Regex.Replace(redacted, "\"(?:Password|UserPassword)\"\\s*:\\s*\"[^\"]*\"", "\"Password\":\"***\"", RegexOptions.IgnoreCase);
        var configuredPassword = _settings.GetTicketSystemPassword();
        return string.IsNullOrEmpty(configuredPassword)
            ? redacted
            : redacted.Replace(configuredPassword, "***", StringComparison.Ordinal);
    }

    private (int created, int updated, int skipped) Fail3(string error)
    {
        LastError = error;
        _logger.Error($"[ZnunyError] {error}");
        return (0, 0, 0);
    }

    public void Dispose()
    {
        _timer.Dispose();
        _syncGate.Dispose();
        _dynamicFieldOptionsGate.Dispose();
        _client.Dispose();
    }

    private sealed class ZnunyApiException : Exception
    {
        public string Stage { get; }
        public HttpStatusCode StatusCode { get; }
        public string ErrorCode { get; }
        public string ErrorMessage { get; }
        public string ResponseBody { get; }

        public ZnunyApiException(string stage, HttpStatusCode statusCode, string errorCode, string errorMessage, string responseBody)
            : base(string.IsNullOrWhiteSpace(errorCode) ? errorMessage : $"{errorCode}: {errorMessage}")
        {
            Stage = stage;
            StatusCode = statusCode;
            ErrorCode = errorCode;
            ErrorMessage = errorMessage;
            ResponseBody = responseBody;
        }
    }

    private sealed class ZnunyTicket
    {
        public string TicketID { get; init; } = string.Empty;
        public string TicketNumber { get; init; } = string.Empty;
        public string Title { get; init; } = string.Empty;
        public string Queue { get; init; } = string.Empty;
        public string State { get; init; } = string.Empty;
        public string StateType { get; init; } = string.Empty;
        public string Priority { get; init; } = string.Empty;
        public string Owner { get; init; } = string.Empty;
        public string Responsible { get; init; } = string.Empty;
        public string Created { get; init; } = string.Empty;
        public string Changed { get; init; } = string.Empty;
        public string DueTime { get; init; } = string.Empty;
        public string PendingTime { get; init; } = string.Empty;
        public string Customer { get; init; } = string.Empty;
        public string CustomerUser { get; init; } = string.Empty;
        public string Lock { get; init; } = string.Empty;
        public string Type { get; init; } = string.Empty;
        public string Service { get; init; } = string.Empty;
        public string SLA { get; init; } = string.Empty;
        public string WebUrl { get; init; } = string.Empty;
        public string DynamicFields { get; init; } = string.Empty;
        public IReadOnlyDictionary<string, string> DynamicFieldValues { get; init; } = new Dictionary<string, string>();
        public string FirstArticleBody { get; init; } = string.Empty;
        public string FirstArticleId { get; init; } = string.Empty;
        public string FirstArticleSenderType { get; init; } = string.Empty;
        public string FirstArticleCreated { get; init; } = string.Empty;
        public int ArticleCount { get; init; }
        public bool IsClosed => IsClosedValue(StateType) || IsClosedValue(State);

        public string GetDynamicFieldValue(string name)
            => string.IsNullOrWhiteSpace(name) || !DynamicFieldValues.TryGetValue(name, out var value) ? string.Empty : value;

        public string FindArticleIdContaining(string marker)
            => Articles.FirstOrDefault(article => article.Body.Contains(marker, StringComparison.OrdinalIgnoreCase))?.ArticleId ?? string.Empty;

        private IReadOnlyList<ZnunyArticle> Articles { get; init; } = Array.Empty<ZnunyArticle>();

        private static bool IsClosedValue(string value)
        {
            var normalized = value.Trim();
            return normalized.Equals("closed", StringComparison.OrdinalIgnoreCase)
                   || normalized.StartsWith("closed ", StringComparison.OrdinalIgnoreCase)
                   || normalized.Equals("removed", StringComparison.OrdinalIgnoreCase)
                   || normalized.StartsWith("removed ", StringComparison.OrdinalIgnoreCase)
                   || normalized.Equals("merged", StringComparison.OrdinalIgnoreCase)
                   || normalized.StartsWith("merged ", StringComparison.OrdinalIgnoreCase);
        }

        public static ZnunyTicket FromJson(JsonElement item, string webBaseUrl, JsonElement? responseRoot = null)
        {
            var id = FirstString(item, "TicketID");
            var number = FirstString(item, "TicketNumber");
            var articles = ExtractArticles(item);
            if (articles.Count == 0 && responseRoot.HasValue)
                articles = ExtractArticles(responseRoot.Value);
            var selectedArticle = articles
                .Where(article => !string.IsNullOrWhiteSpace(article.Body) && !article.IsSystemArticle)
                .OrderBy(article => article.CreatedSort)
                .ThenBy(article => article.ArticleId, StringComparer.OrdinalIgnoreCase)
                .FirstOrDefault();
            return new ZnunyTicket
            {
                TicketID = id,
                TicketNumber = number,
                Title = FirstString(item, "Title"),
                Queue = FirstString(item, "Queue"),
                State = FirstString(item, "State"),
                StateType = FirstString(item, "StateType"),
                Priority = FirstString(item, "Priority"),
                Owner = FirstString(item, "Owner"),
                Responsible = FirstString(item, "Responsible"),
                Created = FirstString(item, "Created", "CreateTime"),
                Changed = FirstString(item, "Changed", "ChangeTime"),
                DueTime = FirstString(item, "DueTime", "EscalationTime"),
                PendingTime = FirstString(item, "PendingTime", "UntilTime"),
                Customer = FirstString(item, "CustomerID", "Customer"),
                CustomerUser = FirstString(item, "CustomerUserID", "CustomerUser"),
                Lock = FirstString(item, "Lock"),
                Type = FirstString(item, "Type"),
                Service = FirstString(item, "Service"),
                SLA = FirstString(item, "SLA"),
                WebUrl = BuildTicketWebUrl(webBaseUrl, id),
                DynamicFields = ExtractDynamicFields(item),
                DynamicFieldValues = ExtractDynamicFieldValues(item),
                Articles = articles,
                ArticleCount = articles.Count,
                FirstArticleBody = selectedArticle?.Body ?? string.Empty,
                FirstArticleId = selectedArticle?.ArticleId ?? string.Empty,
                FirstArticleSenderType = selectedArticle?.SenderType ?? string.Empty,
                FirstArticleCreated = selectedArticle?.Created ?? string.Empty
            };
        }

        public string ToDescription()
        {
            var sb = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(FirstArticleBody))
            {
                sb.AppendLine("--- Erste Ticket-Nachricht ---");
                sb.AppendLine();
                sb.AppendLine(FirstArticleBody);
                sb.AppendLine();
            }
            sb.AppendLine("--- Ticketinformationen ---");
            sb.AppendLine();
            sb.AppendLine($"Znuny TicketID: {TicketID}");
            sb.AppendLine($"TicketNumber: {TicketNumber}");
            sb.AppendLine($"Title: {Title}");
            sb.AppendLine($"Queue: {Queue}");
            sb.AppendLine($"State: {State}");
            if (!string.IsNullOrWhiteSpace(StateType)) sb.AppendLine($"StateType: {StateType}");
            sb.AppendLine($"Priority: {Priority}");
            sb.AppendLine($"Owner: {Owner}");
            sb.AppendLine($"Responsible: {Responsible}");
            sb.AppendLine($"Created: {Created}");
            sb.AppendLine($"Changed: {Changed}");
            sb.AppendLine($"DueTime: {DueTime}");
            sb.AppendLine($"PendingTime: {PendingTime}");
            sb.AppendLine($"Customer: {Customer}");
            sb.AppendLine($"CustomerUser: {CustomerUser}");
            sb.AppendLine($"Lock: {Lock}");
            sb.AppendLine($"Type: {Type}");
            sb.AppendLine($"Service: {Service}");
            sb.AppendLine($"SLA: {SLA}");
            if (!string.IsNullOrWhiteSpace(DynamicFields)) sb.AppendLine($"DynamicFields: {DynamicFields}");
            return sb.ToString();
        }

        private static List<ZnunyArticle> ExtractArticles(JsonElement ticket)
        {
            JsonElement value = default;
            var found = ticket.ValueKind == JsonValueKind.Object
                && ticket.EnumerateObject().Any(property =>
                {
                    if (!string.Equals(property.Name, "Article", StringComparison.OrdinalIgnoreCase)
                        && !string.Equals(property.Name, "Articles", StringComparison.OrdinalIgnoreCase)) return false;
                    value = property.Value;
                    return true;
                });
            if (!found) return new List<ZnunyArticle>();

            var elements = value.ValueKind == JsonValueKind.Array
                ? value.EnumerateArray().ToList()
                : new List<JsonElement> { value };
            return elements
                .Where(element => element.ValueKind == JsonValueKind.Object)
                .Select(element =>
                {
                    var contentType = FirstString(element, "ContentType", "MimeType");
                    var rawBody = FirstString(element, "Body", "BodyPlain", "Content");
                    return new ZnunyArticle(
                        FirstString(element, "ArticleID"),
                        FirstString(element, "SenderType", "SenderTypeID"),
                        FirstString(element, "CreateTime", "Created"),
                        NormalizeArticleBody(rawBody, contentType),
                        FirstString(element, "CommunicationChannel", "ArticleType", "ArticleTypeID"));
                })
                .ToList();
        }

        private static string NormalizeArticleBody(string body, string contentType)
        {
            if (string.IsNullOrWhiteSpace(body)) return string.Empty;
            var text = body;
            if (contentType.Contains("html", StringComparison.OrdinalIgnoreCase) || Regex.IsMatch(text, "<[^>]+>"))
            {
                text = Regex.Replace(text, "<(br|/p|/div|/li|/tr|/h[1-6])[^>]*>", "\n", RegexOptions.IgnoreCase);
                text = Regex.Replace(text, "<li[^>]*>", "- ", RegexOptions.IgnoreCase);
                text = Regex.Replace(text, "<[^>]+>", string.Empty);
                text = WebUtility.HtmlDecode(text);
            }

            text = text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n');
            text = Regex.Replace(text, "[ \t]+\n", "\n");
            text = Regex.Replace(text, "\n{3,}", "\n\n").Trim();
            return text.Length <= 5000 ? text : text[..5000].TrimEnd() + "\n[…]";
        }

        private sealed record ZnunyArticle(string ArticleId, string SenderType, string Created, string Body, string Channel)
        {
            public DateTime CreatedSort => DateTime.TryParse(Created, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var parsed)
                ? parsed
                : DateTime.MaxValue;

            public bool IsSystemArticle
                => SenderType.Contains("system", StringComparison.OrdinalIgnoreCase)
                   || Channel.Contains("system", StringComparison.OrdinalIgnoreCase)
                   || Channel.Contains("internal", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(Body);
        }

        private static string BuildTicketWebUrl(string webBaseUrl, string ticketId)
        {
            if (string.IsNullOrWhiteSpace(webBaseUrl) || string.IsNullOrWhiteSpace(ticketId)) return webBaseUrl;

            var normalizedWebBaseUrl = RemoveOtrsPathSegment(webBaseUrl.Trim());
            var separator = normalizedWebBaseUrl.Contains('?') ? '&' : '?';
            return $"{normalizedWebBaseUrl.TrimEnd('/')}{separator}Action=AgentTicketZoom;TicketID={Uri.EscapeDataString(ticketId)}";
        }

        private static string RemoveOtrsPathSegment(string webBaseUrl)
        {
            if (!Uri.TryCreate(webBaseUrl, UriKind.Absolute, out var uri))
            {
                return Regex.Replace(webBaseUrl, "/otrs(?=/|$)", string.Empty, RegexOptions.IgnoreCase);
            }

            var segments = uri.AbsolutePath
                .Split('/', StringSplitOptions.RemoveEmptyEntries)
                .Where(segment => !string.Equals(segment, "otrs", StringComparison.OrdinalIgnoreCase))
                .ToArray();
            var path = segments.Length == 0 ? "/" : "/" + string.Join("/", segments);

            var builder = new UriBuilder(uri)
            {
                Path = path
            };

            return builder.Uri.ToString();
        }

        private static string ExtractDynamicFields(JsonElement item)
        {
            if (!item.TryGetProperty("DynamicField", out var value)) return string.Empty;
            return value.ToString();
        }

        private static IReadOnlyDictionary<string, string> ExtractDynamicFieldValues(JsonElement item)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (!TryGetPropertyCaseInsensitive(item, "DynamicField", out var value)) return result;
            IEnumerable<JsonElement> fields = value.ValueKind == JsonValueKind.Array
                ? value.EnumerateArray().ToArray()
                : new[] { value };
            foreach (var field in fields)
            {
                if (field.ValueKind != JsonValueKind.Object) continue;
                var name = FirstString(field, "Name");
                if (string.IsNullOrWhiteSpace(name)) continue;
                var fieldValue = FirstString(field, "Value");
                result[name] = fieldValue;
            }
            return result;
        }
    }
}
