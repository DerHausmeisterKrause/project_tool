using System.Net;
using System.Net.Http;
using System.Net.Mail;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Runtime.CompilerServices;
using TaskTool.Models;
using TaskStatus = TaskTool.Models.TaskStatus;

namespace TaskTool.Services;

public class TicketSystemService : IDisposable
{
    private readonly SettingsService _settings;
    private readonly TaskService _tasks;
    private readonly TicketAssignmentSnapshotService _assignmentSnapshots;
    private readonly TicketCandidateSnapshotService _candidateSnapshots;
    private readonly TicketCandidateScanStateService _candidateScanStates;
    private readonly TicketDetailCacheService _detailCache;
    private readonly NotificationService _notifications;
    private readonly LoggerService _logger;
    private readonly HttpClient _client = new() { Timeout = TimeSpan.FromSeconds(45) };
    private readonly SemaphoreSlim _syncGate = new(1, 1);
    private readonly SemaphoreSlim _dynamicFieldOptionsGate = new(1, 1);
    private readonly SemaphoreSlim _sessionGate = new(1, 1);
    private readonly ConcurrentDictionary<string, byte> _manualSelfAssignmentNotificationSuppressions = new(StringComparer.OrdinalIgnoreCase);
    private readonly ConcurrentQueue<ZnunyRequestEvent> _trafficEvents = new();
    private readonly System.Threading.Timer _initialSyncTimer;
    private readonly System.Threading.Timer _fullSyncTimer;
    private readonly System.Threading.Timer _candidateSyncTimer;
    private bool _scheduledSyncStarted;
    private bool _initialSyncCompleted;
    private readonly AsyncLocal<SyncTraffic?> _syncTraffic = new();
    private readonly AsyncLocal<ZnunyRequestReason?> _requestReason = new();
    private IReadOnlyDictionary<string, IReadOnlyList<TicketFieldOption>> _dynamicFieldOptionsCache = new Dictionary<string, IReadOnlyList<TicketFieldOption>>(StringComparer.OrdinalIgnoreCase);
    private DateTime _dynamicFieldOptionsCacheExpiresUtc;
    private bool _dynamicFieldOptionsCacheValid;
    private string _sessionId = string.Empty;
    private DateTime _sessionExpiresUtc;
    private IReadOnlyList<ZnunyCandidateTicket> _candidateTickets = Array.Empty<ZnunyCandidateTicket>();
    private int _lastCandidateUserId;
    private string _lastCandidateKeywords = string.Empty;
    private string _lastCandidateExcludeKeywords = string.Empty;
    private int _lastFullSyncInterval;
    private int _lastCandidateSyncInterval;

    public string LastError { get; private set; } = string.Empty;
    public event Action? TasksChanged;
    public event Action? CandidateTicketsChanged;
    public IReadOnlyList<ZnunyCandidateTicket> CurrentCandidateTickets => _candidateTickets;
    public string CandidateTicketsError { get; private set; } = string.Empty;
    public bool IsCandidateRefreshRunning { get; private set; }
    public string CandidateStatusMessage { get; private set; } = string.Empty;

    private const int MaxIndividualAssignmentNotifications = 5;

    public TicketSystemService(SettingsService settings, TaskService tasks, TicketAssignmentSnapshotService assignmentSnapshots, TicketCandidateSnapshotService candidateSnapshots, TicketCandidateScanStateService candidateScanStates, TicketDetailCacheService detailCache, NotificationService notifications, LoggerService logger)
    {
        _settings = settings;
        _tasks = tasks;
        _assignmentSnapshots = assignmentSnapshots;
        _candidateSnapshots = candidateSnapshots;
        _candidateScanStates = candidateScanStates;
        _detailCache = detailCache;
        _notifications = notifications;
        _logger = logger;
        _initialSyncTimer = new System.Threading.Timer(async _ => await RunInitialSyncAsync(), null, Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
        _fullSyncTimer = new System.Threading.Timer(async _ => await RunFullTimerSyncAsync(), null, Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
        _candidateSyncTimer = new System.Threading.Timer(async _ => await RunCandidateTimerSyncAsync(), null, Timeout.InfiniteTimeSpan, Timeout.InfiniteTimeSpan);
        _lastCandidateUserId = _settings.Current.TicketSystemCandidateUserId;
        _lastCandidateKeywords = _settings.Current.TicketSystemCandidateKeywords;
        _lastCandidateExcludeKeywords = _settings.Current.TicketSystemCandidateExcludeKeywords;
        _lastFullSyncInterval = ZnunySyncPolicy.NormalizeIntervalMinutes(_settings.Current.TicketSystemSyncIntervalMinutes);
        _lastCandidateSyncInterval = ZnunySyncPolicy.NormalizeCandidateIntervalMinutes(_settings.Current.TicketSystemCandidateSyncIntervalMinutes);
        _candidateTickets = _candidateSnapshots.Load();
        LoadPersistedDynamicFieldOptions();
        HandleSettingsChanged();
    }

    public void HandleSettingsChanged()
    {
        var fullInterval = ZnunySyncPolicy.NormalizeIntervalMinutes(_settings.Current.TicketSystemSyncIntervalMinutes);
        var candidateInterval = ZnunySyncPolicy.NormalizeCandidateIntervalMinutes(_settings.Current.TicketSystemCandidateSyncIntervalMinutes);
        var fullIntervalChanged = fullInterval != _lastFullSyncInterval;
        var candidateIntervalChanged = candidateInterval != _lastCandidateSyncInterval;
        var candidateSettingsChanged = _lastCandidateUserId != _settings.Current.TicketSystemCandidateUserId
            || !string.Equals(_lastCandidateKeywords, _settings.Current.TicketSystemCandidateKeywords, StringComparison.Ordinal)
            || !string.Equals(_lastCandidateExcludeKeywords, _settings.Current.TicketSystemCandidateExcludeKeywords, StringComparison.Ordinal);
        _lastCandidateUserId = _settings.Current.TicketSystemCandidateUserId;
        _lastCandidateKeywords = _settings.Current.TicketSystemCandidateKeywords;
        _lastCandidateExcludeKeywords = _settings.Current.TicketSystemCandidateExcludeKeywords;
        _lastFullSyncInterval = fullInterval;
        _lastCandidateSyncInterval = candidateInterval;
        if (candidateSettingsChanged)
        {
            CandidateTicketsError = "Candidate-Einstellungen werden beim nächsten Ticket-Sync angewendet.";
            _candidateScanStates.Replace(Array.Empty<CandidateScanState>());
        }
        if (_scheduledSyncStarted && _initialSyncCompleted)
        {
            if (fullIntervalChanged) ScheduleFullTimer();
            if (candidateIntervalChanged) ScheduleCandidateTimer();
        }
        NotifyTasksChanged();
    }

    public void StartScheduledSync()
    {
        if (_scheduledSyncStarted)
            return;

        _scheduledSyncStarted = true;
        var startupDelay = TimeSpan.FromSeconds(Random.Shared.Next(2, 9));
        _initialSyncTimer.Change(startupDelay, Timeout.InfiniteTimeSpan);
        _logger.Info($"[ZnunyScheduledSync] action=started firstRunSeconds={startupDelay.TotalSeconds:0} fullIntervalMinutes={ZnunySyncPolicy.NormalizeIntervalMinutes(_settings.Current.TicketSystemSyncIntervalMinutes)} candidateIntervalMinutes={ZnunySyncPolicy.NormalizeCandidateIntervalMinutes(_settings.Current.TicketSystemCandidateSyncIntervalMinutes)}");
    }

    private async Task RunInitialSyncAsync()
    {
        if (_initialSyncCompleted) return;
        _logger.Info("[ZnunyScheduledSync] reason=InitialSync pipeline=AssignedThenCandidate");
        if (!await _syncGate.WaitAsync(0))
        {
            var retryDelay = TimeSpan.FromSeconds(Random.Shared.Next(30, 61));
            _logger.Info($"[ZnunySync] reason=InitialSync started=false reasonBusy=true retrySeconds={retryDelay.TotalSeconds:0} initialOnly=true");
            _initialSyncTimer.Change(retryDelay, Timeout.InfiniteTimeSpan);
            return;
        }
        var assignedStarted = false;
        try
        {
            var assigned = await SyncAssignedTicketsAsync(ZnunyRequestReason.InitialSync, gateAlreadyHeld: true);
            assignedStarted = assigned.Started;
            if (assigned.Started)
                await RefreshCandidateTicketsInternalAsync(ZnunyRequestReason.InitialSync, lightweight: false, updateManualStatus: false, gateAlreadyHeld: true);
        }
        catch (Exception ex)
        {
            _logger.Warning($"[ZnunyInitialSync] action=failed message='{LogValue(ex.Message)}'");
        }
        finally
        {
            _syncGate.Release();
            if (assignedStarted)
            {
                _initialSyncCompleted = true;
                ScheduleFullTimer();
                ScheduleCandidateTimer();
            }
        }
    }

    private async Task RunFullTimerSyncAsync()
    {
        await SyncAssignedTicketsAsync(ZnunyRequestReason.FullTimerSync);
        ScheduleFullTimer();
    }

    private async Task RunCandidateTimerSyncAsync()
    {
        await RefreshCandidateTicketsInternalAsync(ZnunyRequestReason.CandidateTimerSync, lightweight: true, updateManualStatus: false);
        ScheduleCandidateTimer();
    }

    private void ScheduleFullTimer()
        => _fullSyncTimer.Change(TimeSpan.FromMinutes(ZnunySyncPolicy.NormalizeIntervalMinutes(_settings.Current.TicketSystemSyncIntervalMinutes)) + TimeSpan.FromSeconds(Random.Shared.Next(0, 21)), Timeout.InfiniteTimeSpan);

    private void ScheduleCandidateTimer()
        => _candidateSyncTimer.Change(TimeSpan.FromMinutes(ZnunySyncPolicy.NormalizeCandidateIntervalMinutes(_settings.Current.TicketSystemCandidateSyncIntervalMinutes)) + TimeSpan.FromSeconds(Random.Shared.Next(0, 21)), Timeout.InfiniteTimeSpan);

    public async Task<ZnunySyncResult> ImportAssignedOpenTicketsAsync()
    {
        var result = await SyncAssignedTicketsAsync(ZnunyRequestReason.ManualFullSync);
        if (_scheduledSyncStarted && _initialSyncCompleted) ScheduleFullTimer();
        return result;
    }

    public async Task<AssignTicketResult> AssignCandidateToCurrentAgentAsync(ZnunyCandidateTicket candidate)
    {
        if (candidate == null || string.IsNullOrWhiteSpace(candidate.TicketId))
            return new AssignTicketResult(false, "Das Ticket besitzt keine gültige TicketID.");

        var configError = ValidateConfiguration(requireAgentId: true);
        if (!string.IsNullOrWhiteSpace(configError))
            return new AssignTicketResult(false, configError);
        if (!await _syncGate.WaitAsync(0))
            return new AssignTicketResult(false, "Es läuft bereits eine Znuny-Aktion. Bitte versuchen Sie es gleich erneut.");

        var ticketId = candidate.TicketId.Trim();
        var targetAgentId = _settings.Current.TicketSystemAgentId;
        var totalStopwatch = Stopwatch.StartNew();
        long sessionMs = 0;
        long updateMs = 0;
        long ticketGetMs = 0;
        long localUpsertMs = 0;
        var serverConfirmed = false;
        var localUpsertSucceeded = false;
        try
        {
            _logger.Info($"[ZnunySelfAssign] ticketId={ticketId} ticketNumber='{LogValue(candidate.TicketNumber)}' targetAgentId={targetAgentId} action=start");
            var stepStopwatch = Stopwatch.StartNew();
            var sessionId = await CreateSessionAsync();
            var sessionHash = HashSessionId(sessionId);
            sessionMs = stepStopwatch.ElapsedMilliseconds;

            stepStopwatch.Restart();
            var route = ResolveTicketUpdateRoute(ticketId);
            var payload = new Dictionary<string, object?>
            {
                ["SessionID"] = sessionId,
                ["TicketID"] = ticketId,
                ["Ticket"] = new Dictionary<string, object?>
                {
                    ["OwnerID"] = targetAgentId,
                    ["ResponsibleID"] = targetAgentId
                }
            };
            using var request = new HttpRequestMessage(HttpMethod.Post, Combine(_settings.Current.TicketSystemApiUrl, route))
            {
                Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
            };
            var response = await SendZnunyAsync(request, "TicketUpdateSelfAssignment", "[ZnunyTicketUpdateResponse]", logBody: false);
            EnsureTicketUpdateResponseIsInterpretable(response);
            updateMs = stepStopwatch.ElapsedMilliseconds;
            serverConfirmed = true;
            _manualSelfAssignmentNotificationSuppressions.TryAdd(ticketId, 0);
            _logger.Info($"[ZnunySelfAssign] ticketId={ticketId} targetAgentId={targetAgentId} action=updated");

            stepStopwatch.Restart();
            try
            {
                localUpsertSucceeded = UpsertAssignedCandidateLocally(candidate);
                localUpsertMs = stepStopwatch.ElapsedMilliseconds;
                RemoveCandidateLocally(ticketId);
                NotifyTasksChanged();
            }
            catch (Exception ex)
            {
                _logger.Error($"[ZnunySelfAssign] ticketId={ticketId} action=local-upsert-failed message='{LogValue(ex.Message)}'");
            }

            _logger.Info($"[ZnunyActionTraffic] reason=ManualAssign ticketId={ticketId} updateRequests=1 metadataRequests=0 fullSyncTriggered=false localUpsert={localUpsertSucceeded.ToString().ToLowerInvariant()}");

            return localUpsertSucceeded
                ? new AssignTicketResult(true, $"Ticket {candidate.TicketNumber} wurde Ihnen zugewiesen.")
                : new AssignTicketResult(true, "Ticket wurde zugewiesen, konnte aber lokal nicht sofort aktualisiert werden. Es erscheint spätestens beim nächsten vollständigen Sync.");
        }
        catch (ZnunyApiException ex) when ((int)ex.StatusCode >= 500)
        {
            LogZnunyError(ex);
            return new AssignTicketResult(false,
                "Zuweisung konnte nicht eindeutig bestätigt werden. Bitte aktualisieren Sie die Aufgaben.",
                ConfirmationUncertain: true);
        }
        catch (ZnunyApiException ex)
        {
            LogZnunyError(ex);
            return new AssignTicketResult(false, $"Ticket konnte nicht zugewiesen werden: {ex.ErrorMessage}");
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or JsonException)
        {
            _logger.Warning($"[ZnunySelfAssign] ticketId={ticketId} targetAgentId={targetAgentId} action=unconfirmed message='{LogValue(ex.Message)}'");
            return new AssignTicketResult(false,
                "Zuweisung konnte nicht eindeutig bestätigt werden. Bitte aktualisieren Sie die Aufgaben.",
                ConfirmationUncertain: true);
        }
        catch (Exception ex)
        {
            _logger.Error($"[ZnunySelfAssign] ticketId={ticketId} targetAgentId={targetAgentId} action=failed message='{LogValue(ex.Message)}'");
            return new AssignTicketResult(false, $"Ticket konnte nicht zugewiesen werden: {ex.Message}");
        }
        finally
        {
            _logger.Info($"[ZnunySelfAssignPerformance] ticketId={ticketId} sessionMs={sessionMs} updateMs={updateMs} ticketGetMs={ticketGetMs} localUpsertMs={localUpsertMs} totalMs={totalStopwatch.ElapsedMilliseconds} serverConfirmed={serverConfirmed.ToString().ToLowerInvariant()} fullSyncTriggered=false");
            _syncGate.Release();
        }
    }

    private bool UpsertAssignedTicketLocally(ZnunyTicket ticket)
    {
        var matchingTasks = _tasks.GetAllTasks()
            .Where(task => string.Equals(ExtractZnunyTicketIdFromTask(task), ticket.TicketID, StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (matchingTasks.Count > 1)
        {
            _logger.Error($"[ZnunySelfAssign] ticketId={ticket.TicketID} action=targeted-upsert-skipped reason=ambiguous-task-mapping count={matchingTasks.Count}");
            return false;
        }

        if (matchingTasks.Count == 0)
        {
            var task = new TaskItem
            {
                IsZnunyAssigned = true,
                Status = ticket.IsClosed ? TaskStatus.Done : TaskStatus.Planned
            };
            MapTicketToTask(ticket, task);
            _tasks.CreateTask(task, isBackgroundImport: true);
            _logger.Info($"[ZnunySelfAssign] ticketId={ticket.TicketID} taskId={task.Id} action=targeted-upsert-created");
            return true;
        }

        var existingTask = matchingTasks[0];
        MapTicketToTask(ticket, existingTask);
        existingTask.IsZnunyAssigned = true;
        if (ticket.IsClosed)
            existingTask.Status = TaskStatus.Done;
        _tasks.UpdateTask(existingTask, touchLocalActivity: false);
        _logger.Info($"[ZnunySelfAssign] ticketId={ticket.TicketID} taskId={existingTask.Id} action=targeted-upsert-updated");
        return true;
    }

    private bool UpsertAssignedCandidateLocally(ZnunyCandidateTicket candidate)
    {
        var matchingTasks = _tasks.GetAllTasks()
            .Where(task => string.Equals(ExtractZnunyTicketIdFromTask(task), candidate.TicketId, StringComparison.OrdinalIgnoreCase)).ToList();
        if (matchingTasks.Count > 1)
        {
            _logger.Error($"[ZnunySelfAssign] ticketId={candidate.TicketId} action=targeted-upsert-skipped reason=ambiguous-task-mapping count={matchingTasks.Count}");
            return false;
        }

        var task = matchingTasks.SingleOrDefault() ?? new TaskItem();
        var cached = _detailCache.Load(candidate.TicketId);
        task.Title = string.IsNullOrWhiteSpace(candidate.Title) ? cached?.TicketTitle ?? candidate.TicketNumber : candidate.Title;
        task.Description = cached?.Articles.FirstOrDefault()?.Body ?? candidate.DescriptionPreview;
        task.TicketUrl = candidate.WebUrl;
        task.Tags = AddZnunyTicketTags(task.Tags, candidate.TicketId, candidate.TicketNumber);
        task.TicketState = candidate.State;
        task.IsZnunyAssigned = true;
        if (task.Status != TaskStatus.Done) task.Status = TaskStatus.Planned;
        if (matchingTasks.Count == 0) _tasks.CreateTask(task, isBackgroundImport: true);
        else _tasks.UpdateTask(task, touchLocalActivity: false);
        _logger.Info($"[ZnunySelfAssign] ticketId={candidate.TicketId} taskId={task.Id} action=targeted-upsert-{(matchingTasks.Count == 0 ? "created" : "updated")} cacheUsed={(cached != null).ToString().ToLowerInvariant()}");
        return true;
    }

    private void RemoveCandidateLocally(string ticketId)
    {
        var remaining = _candidateTickets
            .Where(candidate => !string.Equals(candidate.TicketId, ticketId, StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (remaining.Count == _candidateTickets.Count) return;
        _candidateSnapshots.Replace(remaining);
        _candidateScanStates.Remove(ticketId);
        PublishCandidateTickets(remaining, string.Empty);
        _logger.Info($"[ZnunySelfAssign] ticketId={ticketId} action=candidate-removed-locally");
    }

    public async Task<CreateTicketResult> CreateTicketFromLocalTaskAsync(TaskItem task)
    {
        if (task == null)
            return new CreateTicketResult(false, "Es wurde keine Aufgabe ausgewählt.");
        if (task.IsZnunyTask)
            return new CreateTicketResult(false, "Diese Aufgabe ist bereits mit einem Znuny-Ticket verknüpft.");

        var validationError = ValidateTicketCreateConfiguration(task);
        if (!string.IsNullOrWhiteSpace(validationError))
            return new CreateTicketResult(false, validationError);
        if (!await _syncGate.WaitAsync(0))
            return new CreateTicketResult(false, "Es läuft bereits eine Znuny-Aktion. Bitte versuchen Sie es gleich erneut.");

        try
        {
            _logger.Info($"[ZnunyTicketCreate] taskId={task.Id} action=start");
            var sessionId = await CreateSessionAsync();
            var payload = BuildTicketCreatePayload(task, sessionId);
            var route = _settings.Current.TicketSystemTicketCreateRoute;
            var method = new HttpMethod(_settings.Current.TicketSystemTicketCreateMethod);
            using var request = new HttpRequestMessage(method, Combine(_settings.Current.TicketSystemApiUrl, route))
            {
                Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
            };

            var response = await SendZnunyAsync(request, "TicketCreate", "[ZnunyTicketCreateResponse]", logBody: false);
            using var document = JsonDocument.Parse(response.Body);
            ThrowIfApiError(document.RootElement, "TicketCreate", response.StatusCode, response.Body);
            var ticketId = FindStringRecursive(document.RootElement, "TicketID");
            var ticketNumber = FindStringRecursive(document.RootElement, "TicketNumber");
            if (string.IsNullOrWhiteSpace(ticketId) || string.IsNullOrWhiteSpace(ticketNumber))
            {
                _logger.Warning($"[ZnunyTicketCreate] taskId={task.Id} action=unconfirmed reason=missing-ticket-identifiers");
                return new CreateTicketResult(false,
                    "Ticket-Erstellung konnte nicht eindeutig bestätigt werden. Bitte vor einem erneuten Versuch im Ticketsystem prüfen.",
                    ConfirmationUncertain: true);
            }

            task.Tags = AddZnunyTicketTags(task.Tags, ticketId, ticketNumber);
            task.TicketUrl = BuildTicketWebUrl(_settings.Current.TicketSystemWebUrl, ticketId);
            task.IsZnunyAssigned = true;
            _tasks.UpdateTask(task);
            NotifyTasksChanged();

            _logger.Info($"[ZnunyTicketCreate] taskId={task.Id} ticketId={ticketId} ticketNumber='{LogValue(ticketNumber)}' action=completed");
            return new CreateTicketResult(true, $"Ticket {ticketNumber} wurde erfolgreich erstellt.", ticketId, ticketNumber, task.TicketUrl);
        }
        catch (TaskCanceledException ex)
        {
            _logger.Warning($"[ZnunyTicketCreate] taskId={task.Id} action=unconfirmed message='{LogValue(ex.Message)}'");
            return new CreateTicketResult(false,
                "Ticket-Erstellung konnte nicht eindeutig bestätigt werden. Bitte vor einem erneuten Versuch im Ticketsystem prüfen.",
                ConfirmationUncertain: true);
        }
        catch (HttpRequestException ex)
        {
            _logger.Warning($"[ZnunyTicketCreate] taskId={task.Id} action=unconfirmed message='{LogValue(ex.Message)}'");
            return new CreateTicketResult(false,
                "Ticket-Erstellung konnte nicht eindeutig bestätigt werden. Bitte vor einem erneuten Versuch im Ticketsystem prüfen.",
                ConfirmationUncertain: true);
        }
        catch (ZnunyApiException ex)
        {
            LogZnunyError(ex);
            _logger.Error($"[ZnunyTicketCreate] taskId={task.Id} action=failed errorCode='{LogValue(ex.ErrorCode)}' message='{LogValue(ex.ErrorMessage)}'");
            return new CreateTicketResult(false, $"Ticket konnte nicht erstellt werden: {ex.ErrorMessage}");
        }
        catch (Exception ex)
        {
            _logger.Error($"[ZnunyTicketCreate] taskId={task.Id} action=failed message='{LogValue(ex.Message)}'");
            return new CreateTicketResult(false, $"Ticket konnte nicht erstellt werden: {ex.Message}");
        }
        finally
        {
            _syncGate.Release();
        }
    }

    public async Task<bool> RefreshCandidateTicketsAsync()
        => await RefreshCandidateTicketsInternalAsync(ZnunyRequestReason.ManualCandidateRefresh, lightweight: false, updateManualStatus: true);

    private async Task<bool> RefreshCandidateTicketsInternalAsync(ZnunyRequestReason reason, bool lightweight, bool updateManualStatus, bool gateAlreadyHeld = false)
    {
        var ownsGate = !gateAlreadyHeld;
        if (ownsGate && !await _syncGate.WaitAsync(0))
        {
            _logger.Info($"[ZnunyActionTraffic] reason={reason} skipped=true cause=sync-active requests=0 retry=false");
            if (updateManualStatus)
            {
                CandidateStatusMessage = "Aktualisierung nicht gestartet – derzeit läuft eine andere Znuny-Aktion.";
                CandidateTicketsChanged?.Invoke();
            }
            return false;
        }
        if (updateManualStatus)
        {
            IsCandidateRefreshRunning = true;
            CandidateTicketsError = string.Empty;
            CandidateStatusMessage = "Neue Aufgaben werden aktualisiert …";
            CandidateTicketsChanged?.Invoke();
        }
        try
        {
            LastError = string.Empty;
            var configError = ValidateConfiguration(requireAgentId: false);
            if (!string.IsNullOrWhiteSpace(configError))
            {
                LastError = configError;
                if (updateManualStatus) CandidateStatusMessage = "Aktualisierung fehlgeschlagen – letzter erfolgreicher Stand wird angezeigt.";
                return false;
            }
            var traffic = new SyncTraffic(reason, Stopwatch.StartNew());
            _syncTraffic.Value = traffic;
            try
            {
                var sessionId = await CreateSessionAsync();
                var discovery = await RefreshCandidateTicketsCoreAsync(sessionId, HashSessionId(sessionId), reason, lightweight);
                if (updateManualStatus)
                    CandidateStatusMessage = $"Aktualisiert um {DateTime.Now:HH:mm} · {(discovery.Matched == 0 ? "keine neuen Aufgaben" : $"{discovery.Matched} neue Aufgaben")}";
                _logger.Info($"[ZnunyActionTraffic] reason={reason} searches={traffic.SearchRequests} newCandidateIds={discovery.NewIds} existingCandidateIds={discovery.ExistingIds} candidateDetailGets={traffic.CandidateDetailGets}");
                return true;
            }
            finally { LogTraffic(traffic); _syncTraffic.Value = null; }
        }
        catch (Exception ex)
        {
            LastError = $"Neue Aufgaben konnten nicht aktualisiert werden: {ex.Message}";
            CandidateTicketsError = "Aktualisierung fehlgeschlagen – letzter erfolgreicher Stand wird angezeigt.";
            CandidateStatusMessage = CandidateTicketsError;
            return false;
        }
        finally
        {
            if (updateManualStatus) IsCandidateRefreshRunning = false;
            CandidateTicketsChanged?.Invoke();
            if (ownsGate) _syncGate.Release();
        }
    }

    public async Task<TicketBookingContext> GetTicketBookingContextAsync(TaskItem task)
    {
        var ticketId = ExtractZnunyTicketIdFromTask(task);
        if (string.IsNullOrWhiteSpace(ticketId))
            throw new InvalidOperationException("Der ausgewählte Task besitzt keine eindeutige Znuny-TicketID.");

        var cached = _detailCache.Load(ticketId);
        if (cached != null) return cached;
        await Task.CompletedTask;
        throw new InvalidOperationException("Ticketdetails werden beim nächsten Ticket-Sync geladen.");
    }

    public async Task<TicketBookingContext> RefreshTicketBookingContextAsync(TaskItem task)
    {
        var previousTraffic = _syncTraffic.Value;
        var traffic = new SyncTraffic(ZnunyRequestReason.ManualTicketRefresh, Stopwatch.StartNew());
        _syncTraffic.Value = traffic;
        try
        {
        var ticketId = ExtractZnunyTicketIdFromTask(task);
        if (string.IsNullOrWhiteSpace(ticketId)) throw new InvalidOperationException("Der ausgewählte Task besitzt keine eindeutige Znuny-TicketID.");
        var configError = ValidateConfiguration(requireAgentId: false);
        if (!string.IsNullOrWhiteSpace(configError)) throw new InvalidOperationException(configError);
        var sessionId = await CreateSessionAsync();
        var ticket = await GetTicketDetailsAsync(ticketId, sessionId, HashSessionId(sessionId))
                     ?? throw new InvalidOperationException("TicketGet lieferte keine Ticketdaten.");
        var costField = _settings.Current.TicketSystemCostCenterFieldName;
        var orderField = _settings.Current.TicketSystemOrderFieldName;
        var optionFields = _dynamicFieldOptionsCacheValid
            ? _dynamicFieldOptionsCache
            : new Dictionary<string, IReadOnlyList<TicketFieldOption>>(StringComparer.OrdinalIgnoreCase);
        var costOptions = GetFieldOptions(optionFields, costField, _settings.Current.TicketSystemCostCenterOptions);
        var orderOptions = GetFieldOptions(optionFields, orderField, _settings.Current.TicketSystemOrderOptions);
        var costCenterValue = ticket.GetDynamicFieldValue(costField);
        var orderValue = ticket.GetDynamicFieldValue(orderField);
        _logger.Info($"[ZnunyTicketDynamicFields] ticketId={ticket.TicketID} requestedDynamicFields=true availableFields=[{string.Join(',', ticket.DynamicFieldValues.Keys.OrderBy(name => name, StringComparer.OrdinalIgnoreCase))}] costCenterValue='{costCenterValue}' orderValue='{orderValue}'");
        var fieldsMissing = !ticket.DynamicFieldValues.ContainsKey(costField) || !ticket.DynamicFieldValues.ContainsKey(orderField);
        var information = fieldsMissing ? "Kostenstelle/Auftrag konnten nicht aus OTRS geladen werden." : string.Empty;
        LogDynamicFieldSelection(ticket.TicketID, costField, costCenterValue, costOptions);
        LogDynamicFieldSelection(ticket.TicketID, orderField, orderValue, orderOptions);
        var articles = ticket.ToArticleItems();
        var (replySource, replyRecipient) = ResolveReplyRecipient(articles, ticket.CustomerUser);
        _logger.Info($"[ZnunyConversation] ticketId={ticket.TicketID} articleCount={articles.Count} selectedArticleId='{articles.FirstOrDefault()?.ArticleId ?? string.Empty}'");

        var context = new TicketBookingContext(
            ticket.TicketID,
            ticket.TicketNumber,
            costCenterValue,
            orderValue,
            costOptions,
            orderOptions,
            information,
            articles,
            replySource,
            replyRecipient,
            ticket.Title);
        _detailCache.Store(context, ticket.State, ParseZnunyUtc(ticket.Changed));
        return context;
        }
        finally { LogTraffic(traffic); _syncTraffic.Value = previousTraffic; }
    }

    public async Task<TicketBookingContext> RefreshDynamicFieldOptionsAsync(TaskItem task)
    {
        var previousTraffic = _syncTraffic.Value;
        var traffic = new SyncTraffic(ZnunyRequestReason.ManualDynamicFieldRefresh, Stopwatch.StartNew());
        _syncTraffic.Value = traffic;
        try
        {
            var sessionId = await CreateSessionAsync();
            var fields = await GetDynamicFieldOptionsAsync(sessionId, forceRefresh: true);
            var cached = await GetTicketBookingContextAsync(task);
            return cached with
            {
                CostCenterOptions = GetFieldOptions(fields, _settings.Current.TicketSystemCostCenterFieldName, _settings.Current.TicketSystemCostCenterOptions),
                OrderOptions = GetFieldOptions(fields, _settings.Current.TicketSystemOrderFieldName, _settings.Current.TicketSystemOrderOptions)
            };
        }
        finally { LogTraffic(traffic); _syncTraffic.Value = previousTraffic; }
    }

    public async Task<TicketReplyResult> SendTicketReplyAsync(
        TaskItem task,
        TicketArticleItem? originalCustomerArticle,
        string recipient,
        string ticketTitle,
        string replyText)
    {
        var ticketId = ExtractZnunyTicketIdFromTask(task);
        if (string.IsNullOrWhiteSpace(ticketId))
            return new TicketReplyResult(false, "Der ausgewählte Task besitzt keine eindeutige Znuny-TicketID.");
        var configError = ValidateConfiguration(requireAgentId: false);
        if (!string.IsNullOrWhiteSpace(configError))
            return new TicketReplyResult(false, configError);
        if (!TryNormalizeMailAddress(recipient, out var safeRecipient))
            return new TicketReplyResult(false, "Für dieses Ticket konnte keine eindeutige Empfängeradresse ermittelt werden. Bitte antworten Sie über Znuny.");
        var body = (replyText ?? string.Empty).Trim();
        if (body.Length == 0)
            return new TicketReplyResult(false, "Bitte geben Sie einen Antworttext ein.");
        if (body.Length > 10000)
            body = body[..10000];
        var htmlBody = ConvertPlainTextReplyToHtml(body);
        if (!await _syncGate.WaitAsync(0))
            return new TicketReplyResult(false, "Es läuft bereits eine Znuny-Aktion. Bitte versuchen Sie es gleich erneut.");

        try
        {
            _logger.Info($"[ZnunyReply] ticketId={ticketId} recipientResolved=true recipientFormat=bare-address replyLength={body.Length} mimeType=text/html appendSignature=true action=send");
            var sessionId = await CreateSessionAsync();
            var originalSubject = originalCustomerArticle?.Subject ?? string.Empty;
            var subjectSource = string.IsNullOrWhiteSpace(originalSubject) ? ticketTitle : originalSubject;
            var subject = Regex.IsMatch(subjectSource, @"^\s*re\s*:", RegexOptions.IgnoreCase)
                ? subjectSource.Trim()
                : $"Re: {subjectSource.Trim()}";
            if (string.Equals(subject, "Re: ", StringComparison.Ordinal))
                subject = "Re: Ticketanfrage";
            var payload = new Dictionary<string, object?>
            {
                ["SessionID"] = sessionId,
                ["TicketID"] = ticketId,
                ["Article"] = new Dictionary<string, object?>
                {
                    ["CommunicationChannel"] = "Email",
                    ["SenderType"] = "agent",
                    ["IsVisibleForCustomer"] = 1,
                    ["ArticleSend"] = 1,
                    ["AppendSignatureToBody"] = 1,
                    ["To"] = safeRecipient,
                    ["Subject"] = subject,
                    ["Body"] = htmlBody,
                    ["ContentType"] = "text/html; charset=utf-8",
                    ["MimeType"] = "text/html",
                    ["Charset"] = "utf-8"
                }
            };
            var route = ResolveTicketUpdateRoute(ticketId);
            using var request = new HttpRequestMessage(HttpMethod.Post, Combine(_settings.Current.TicketSystemApiUrl, route))
            {
                Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
            };
            var response = await SendZnunyAsync(request, "TicketUpdateReply", "[ZnunyTicketUpdateResponse]", logBody: false);
            EnsureTicketUpdateResponseIsInterpretable(response);
            var articleId = ExtractFirstValueRecursive(response.Body, "ArticleID");
            _tasks.TouchTaskActivity(task.Id);
            _logger.Info($"[ZnunyReply] ticketId={ticketId} articleId='{articleId}' action=completed");
            return new TicketReplyResult(true, "Antwort wurde gesendet.", articleId);
        }
        catch (ZnunyApiException ex) when ((int)ex.StatusCode >= 500)
        {
            LogZnunyError(ex);
            return new TicketReplyResult(false, "Der Versand konnte nicht eindeutig bestätigt werden. Bitte prüfen Sie den Ticketverlauf, bevor Sie erneut senden.", ConfirmationUncertain: true);
        }
        catch (ZnunyApiException ex)
        {
            LogZnunyError(ex);
            return new TicketReplyResult(false, $"Antwort konnte nicht gesendet werden: {ex.ErrorMessage}");
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or JsonException)
        {
            _logger.Warning($"[ZnunyReply] ticketId={ticketId} replyLength={body.Length} action=unconfirmed message='{LogValue(ex.Message)}'");
            return new TicketReplyResult(false, "Der Versand konnte nicht eindeutig bestätigt werden. Bitte prüfen Sie den Ticketverlauf, bevor Sie erneut senden.", ConfirmationUncertain: true);
        }
        catch (Exception ex)
        {
            _logger.Error($"[ZnunyReply] ticketId={ticketId} replyLength={body.Length} action=failed message='{LogValue(ex.Message)}'");
            return new TicketReplyResult(false, $"Antwort konnte nicht gesendet werden: {ex.Message}");
        }
        finally
        {
            _syncGate.Release();
        }
    }

    private static (TicketArticleItem? source, string recipient) ResolveReplyRecipient(
        IReadOnlyList<TicketArticleItem> articles,
        string customerUser)
    {
        var customerArticles = articles
            .Where(article => article.SenderType.Contains("customer", StringComparison.OrdinalIgnoreCase))
            .ToList();
        var source = customerArticles.FirstOrDefault(article => article.CommunicationChannel.Contains("email", StringComparison.OrdinalIgnoreCase))
                     ?? customerArticles.FirstOrDefault();
        if (source != null && TryNormalizeMailAddress(source.ReplyTo, out var replyTo)) return (source, replyTo);
        if (source != null && TryNormalizeMailAddress(source.From, out var from)) return (source, from);
        return TryNormalizeMailAddress(customerUser, out var fallback) ? (source, fallback) : (source, string.Empty);
    }

    private static bool TryNormalizeMailAddress(string? value, out string normalized)
    {
        normalized = string.Empty;
        if (string.IsNullOrWhiteSpace(value))
            return false;
        try
        {
            var address = new MailAddress(value.Trim());
            if (!address.Address.Contains("@", StringComparison.Ordinal) || address.Address.EndsWith("@", StringComparison.Ordinal)) return false;
            normalized = address.Address;
            return true;
        }
        catch (FormatException)
        {
            return false;
        }
    }

    private static string ConvertPlainTextReplyToHtml(string text)
    {
        var normalized = (text ?? string.Empty)
            .Replace("\r\n", "\n", StringComparison.Ordinal)
            .Replace('\r', '\n');
        var encoded = WebUtility.HtmlEncode(normalized);
        return encoded.Replace("\n", "<br />\n", StringComparison.Ordinal) + "<br /><br />";
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
        string? note,
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
                var currentTicket = await GetTicketDetailsAsync(ticketId, sessionId, sessionHash);
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
                Note = NormalizeBookingNote(note),
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
            _logger.Info($"[ZnunyTimeBooking] route={route} ticketId={ticketId} taskId={task.Id} bookingId={booking.BookingId} minutes={minutes:0.##} timeUnit={timeUnit:0.####} notePresent={!string.IsNullOrEmpty(booking.Note)} noteLength={booking.Note.Length} action=send");
            var response = await SendZnunyAsync(request, "TicketUpdateTimeBooking", "[ZnunyTicketUpdateResponse]", logBody: false);
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
            _logger.Error($"[ZnunyTimeBooking] ticketId={ticketId} taskId={task.Id} action=api-failed httpStatus={(int)ex.StatusCode} errorCode={LogValue(ex.ErrorCode)} message={LogValue(ex.ErrorMessage)}");
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
            var ticket = await GetTicketReconciliationDetailsAsync(booking.TicketId, sessionId, HashSessionId(sessionId));
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
            var ticket = await GetTicketReconciliationDetailsAsync(booking.TicketId, sessionId, sessionHash);
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
            var response = await SendZnunyAsync(request, "TicketUpdateTimeBookingRetry", "[ZnunyTicketUpdateResponse]", logBody: false);
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
                var ticket = await GetTicketDetailsAsync(uniqueIds[0], sessionId, sessionHash);
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

    public async Task<ZnunySyncResult> SyncAssignedTicketsAsync(ZnunyRequestReason reason, bool gateAlreadyHeld = false)
    {
        var ownsGate = !gateAlreadyHeld;
        var gateAcquired = gateAlreadyHeld || (reason == ZnunyRequestReason.ManualFullSync
            ? await _syncGate.WaitAsync(TimeSpan.FromSeconds(8))
            : await _syncGate.WaitAsync(0));
        if (!gateAcquired)
        {
            _logger.Info($"[ZnunySync] reason={reason} started=false reasonBusy=true");
            return ZnunySyncResult.BusyResult();
        }

        LastError = string.Empty;
        var created = 0;
        var updated = 0;
        var unchanged = 0;
        var skipped = 0;
        var searchResultCount = 0;
        var uniqueTicketCount = 0;
        var searchLimitReached = false;
        var hasTaskChanges = false;
        var traffic = new SyncTraffic(reason, Stopwatch.StartNew());
        _syncTraffic.Value = traffic;

        try
        {
            var configError = ValidateConfiguration(requireAgentId: true);
            if (!string.IsNullOrWhiteSpace(configError))
                return ZnunySyncResult.Failed(configError);
            if (!_settings.Current.TicketSystemIncludeOwner && !_settings.Current.TicketSystemIncludeResponsible)
                return ZnunySyncResult.Failed("Znuny Sync benötigt Owner oder Responsible als Suchkriterium.");

            var userId = GetConfiguredAgentId()!.Value;
            var sessionId = await CreateSessionAsync();
            var sessionHash = HashSessionId(sessionId);
            _logger.Info($"[ZnunySyncStart] reason={reason} started=true agentId={userId} searchLimit={EffectiveSearchLimit} sortBy={ZnunySyncPolicy.SearchSortBy} orderBy={ZnunySyncPolicy.SearchOrderBy} baseUrl='{SanitizeUrl(_settings.Current.TicketSystemApiUrl)}' auth={_settings.Current.TicketSystemTicketSearchAuthMode} onlyOpen={_settings.Current.TicketSystemOnlyOpenTickets} showClosed={_settings.Current.TicketSystemShowClosedTickets} includeOwner={_settings.Current.TicketSystemIncludeOwner} includeResponsible={_settings.Current.TicketSystemIncludeResponsible}");
            _logger.Info($"[ZnunyUser] source=ConfiguredSettings userId={userId}");

            var (ownerIds, responsibleIds, uniqueTicketIds) = await SearchAssignedTicketIdsAsync(
                userId,
                sessionId,
                sessionHash,
                _settings.Current.TicketSystemIncludeOwner,
                _settings.Current.TicketSystemIncludeResponsible);
            searchResultCount = ownerIds.Count + responsibleIds.Count;
            uniqueTicketCount = uniqueTicketIds.Count;
            searchLimitReached = ownerIds.Count >= EffectiveSearchLimit || responsibleIds.Count >= EffectiveSearchLimit;
            _logger.Info($"[ZnunyAssignedDiscovery] agentId={userId} owner={ownerIds.Count} responsible={responsibleIds.Count} unique={uniqueTicketIds.Count}");
            var assignmentContextKey = BuildAssignmentContextKey(userId);
            var assignmentContextHash = assignmentContextKey[..12];
            var previousAssignmentSnapshot = _assignmentSnapshots.Load(assignmentContextKey);
            var currentAssignedIds = uniqueTicketIds.ToHashSet(StringComparer.OrdinalIgnoreCase);
            var newlyAssignedIds = previousAssignmentSnapshot.Initialized
                ? currentAssignedIds.Except(previousAssignmentSnapshot.TicketIds, StringComparer.OrdinalIgnoreCase).ToHashSet(StringComparer.OrdinalIgnoreCase)
                : new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var removedAssignmentCount = previousAssignmentSnapshot.Initialized
                ? previousAssignmentSnapshot.TicketIds.Except(currentAssignedIds, StringComparer.OrdinalIgnoreCase).Count()
                : 0;
            _logger.Info($"[ZnunyAssignmentState] previous={previousAssignmentSnapshot.TicketIds.Count} current={currentAssignedIds.Count} new={newlyAssignedIds.Count} removed={removedAssignmentCount}");
            var notificationCandidates = new Dictionary<string, (Guid TaskId, string TicketNumber, string TicketTitle)>(StringComparer.OrdinalIgnoreCase);
            var assignedTicketsFullyProcessed = true;
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

            var previouslyAssignedNowMissing = previousAssignmentSnapshot.Initialized
                ? previousAssignmentSnapshot.TicketIds.Except(currentAssignedIds, StringComparer.OrdinalIgnoreCase)
                : Array.Empty<string>();
            var ticketIds = ZnunySyncPolicy.SelectTicketIds(uniqueTicketIds, previouslyAssignedNowMissing);
            var duplicateCount = ownerIds.Count + responsibleIds.Count - uniqueTicketIds.Count;
            _logger.Info($"[ZnunySearchMerge] ownerTickets={ownerIds.Count} responsibleTickets={responsibleIds.Count} uniqueTickets={uniqueTicketIds.Count} limit={EffectiveSearchLimit} sortBy={ZnunySyncPolicy.SearchSortBy} orderBy={ZnunySyncPolicy.SearchOrderBy} existingTicketsRechecked={ticketIds.Count - uniqueTicketIds.Count} duplicateOwnerResponsibleTickets={duplicateCount}");

            foreach (var ticketId in ticketIds)
            {
                var ticket = await GetTicketMetadataAsync(ticketId, sessionId, sessionHash);
                if (ticket == null)
                {
                    skipped++;
                    if (currentAssignedIds.Contains(ticketId)) assignedTicketsFullyProcessed = false;
                    continue;
                }

                var remoteChangedUtc = ParseZnunyUtc(ticket.Changed);
                var cachedChangedUtc = _detailCache.GetRemoteChangedUtc(ticket.TicketID);
                if (!cachedChangedUtc.HasValue || (remoteChangedUtc.HasValue && remoteChangedUtc > cachedChangedUtc))
                {
                    var details = await GetTicketDetailsAsync(ticketId, sessionId, sessionHash);
                    if (details == null)
                        throw new InvalidOperationException($"TicketGet lieferte keine Detaildaten für TicketID {ticketId}; der Assignment-Snapshot bleibt unverändert.");
                    ticket = details;
                    StoreTicketDetails(details);
                    traffic.DetailRemote++;
                }
                else
                {
                    traffic.DetailCacheHits++;
                }

                if (ambiguousTicketIds.Contains(ticket.TicketID))
                {
                    skipped++;
                    continue;
                }

                var isCurrentlyAssigned = currentAssignedIds.Contains(ticket.TicketID);
                if (existing.TryGetValue(ticket.TicketID, out var task))
                {
                    var taskChanged = false;
                    if (!isCurrentlyAssigned && task.Status == TaskStatus.Running)
                        _tasks.StopTimer(task);
                    var metadataChanged = MapTicketToTask(ticket, task);
                    if (ticket.IsClosed)
                    {
                        if (task.Status == TaskStatus.Done)
                        {
                            if (metadataChanged)
                            {
                                _tasks.UpdateTask(task, touchLocalActivity: false);
                                updated++;
                                taskChanged = true;
                                hasTaskChanges = true;
                            }

                            LogAutoComplete(ticket, task, "already-completed");
                        }
                        else
                        {
                            // Use the same persistence path as the manual "Erledigt" action.
                            task.Status = TaskStatus.Done;
                            _tasks.UpdateTask(task, touchLocalActivity: false);
                            updated++;
                            taskChanged = true;
                            hasTaskChanges = true;
                            LogAutoComplete(ticket, task, "completed");
                        }
                    }
                    else if (metadataChanged)
                    {
                        // Deliberately preserve Done: reopening in Znuny is not propagated back.
                        _tasks.UpdateTask(task, touchLocalActivity: false);
                        updated++;
                        taskChanged = true;
                        hasTaskChanges = true;
                        _logger.Info($"[ZnunyTaskUpdated] ticketId={ticket.TicketID} ticketNumber='{ticket.TicketNumber}' taskId={task.Id}");
                    }

                    if (task.IsZnunyAssigned != isCurrentlyAssigned)
                    {
                        _tasks.SetZnunyAssigned(task, isCurrentlyAssigned);
                        hasTaskChanges = true;
                        if (!taskChanged) updated++;
                        taskChanged = true;
                        _logger.Info($"[ZnunyAssignmentState] ticketId={ticket.TicketID} taskId={task.Id} assigned={isCurrentlyAssigned.ToString().ToLowerInvariant()} action={(isCurrentlyAssigned ? "show" : "hide")}");
                    }
                    if (!taskChanged) unchanged++;
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
                    task.IsZnunyAssigned = true;
                    task.Status = ticket.IsClosed ? TaskStatus.Done : TaskStatus.Planned;
                    _tasks.CreateTask(task, isBackgroundImport: true);
                    created++;
                    hasTaskChanges = true;
                    _logger.Info($"[ZnunyTaskCreated] ticketId={ticket.TicketID} ticketNumber='{ticket.TicketNumber}' taskId={task.Id}");
                }

                if (!ticket.IsClosed && newlyAssignedIds.Contains(ticket.TicketID))
                {
                    if (_manualSelfAssignmentNotificationSuppressions.ContainsKey(ticket.TicketID))
                    {
                        _logger.Info($"[ZnunySelfAssign] ticketId={ticket.TicketID} action=notification-suppressed snapshotPending=true");
                    }
                    else
                    {
                        notificationCandidates[ticket.TicketID] = (task.Id, ticket.TicketNumber, ticket.Title);
                        _logger.Info($"[ZnunyNotificationCandidate] ticketId={ticket.TicketID} ticketNumber='{ticket.TicketNumber}' taskId={task.Id}");
                    }
                }
            }

            foreach (var noLongerAssigned in existing.Where(item => !currentAssignedIds.Contains(item.Key)))
            {
                var task = noLongerAssigned.Value;
                if (task.Status == TaskStatus.Running)
                    _tasks.StopTimer(task);
                if (!task.IsZnunyAssigned) continue;
                _tasks.SetZnunyAssigned(task, false);
                hasTaskChanges = true;
                _logger.Info($"[ZnunyAssignmentState] ticketId={noLongerAssigned.Key} taskId={task.Id} assigned=false action=hide");
            }

            if (!assignedTicketsFullyProcessed)
                throw new InvalidOperationException("Mindestens ein aktuell zugewiesenes Ticket konnte nicht vollständig verarbeitet werden; der Assignment-Snapshot bleibt unverändert.");

            if (!previousAssignmentSnapshot.Initialized)
            {
                _assignmentSnapshots.Replace(assignmentContextKey, currentAssignedIds);
                _logger.Info($"[ZnunyAssignmentSnapshot] committed=true contextHash={assignmentContextHash} current={currentAssignedIds.Count}");
                CompleteManualSelfAssignmentSuppressions(currentAssignedIds);
                _logger.Info($"[ZnunyAssignmentNotifications] action=initialize-baseline contextHash={assignmentContextHash} current={currentAssignedIds.Count} notifications=0");
            }
            else
            {
                var notificationsEnabled = _settings.Current.NotifyOnNewAssignedTickets;
                var notificationCount = notificationsEnabled
                    ? notificationCandidates.Count > MaxIndividualAssignmentNotifications ? 1 : notificationCandidates.Count
                    : 0;
                _logger.Info($"[ZnunyAssignmentNotifications] contextHash={assignmentContextHash} previous={previousAssignmentSnapshot.TicketIds.Count} current={currentAssignedIds.Count} newlyAssigned={newlyAssignedIds.Count} removed={removedAssignmentCount} notifications={notificationCount}");

                if (notificationsEnabled && notificationCandidates.Count > 0)
                {
                    IReadOnlyList<TicketNotificationPayload> payloads;
                    if (notificationCandidates.Count > MaxIndividualAssignmentNotifications)
                    {
                        payloads = new[]
                        {
                            new TicketNotificationPayload(Guid.Empty, $"Du hast {notificationCandidates.Count} neue Tickets\nÖffne Plenaro, um die neuen Aufgaben anzusehen.")
                        };
                        _logger.Info($"[ZnunyAssignmentNotifications] newlyAssigned={notificationCandidates.Count} mode=summary individualNotifications=0");
                    }
                    else
                    {
                        payloads = notificationCandidates
                            .OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase)
                            .Select(candidate => new TicketNotificationPayload(
                                candidate.Value.TaskId,
                                $"Du hast ein neues Ticket\n{candidate.Value.TicketNumber} · {candidate.Value.TicketTitle}"))
                            .ToList();
                    }

                    if (!await _notifications.EnqueueTicketNotificationsAsync(payloads))
                        throw new InvalidOperationException("Die Ticket-Benachrichtigungen wurden nicht von der Dynamic-Island-Queue angenommen; der Assignment-Snapshot bleibt unverändert.");

                    foreach (var candidate in notificationCandidates.OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase))
                        _logger.Info($"[ZnunyNewAssignment] ticketId={candidate.Key} ticketNumber='{candidate.Value.TicketNumber}' taskId={candidate.Value.TaskId} action=notify");
                }

                _assignmentSnapshots.Replace(assignmentContextKey, currentAssignedIds);
                _logger.Info($"[ZnunyAssignmentSnapshot] committed=true contextHash={assignmentContextHash} current={currentAssignedIds.Count}");
                CompleteManualSelfAssignmentSuppressions(currentAssignedIds);
            }

            _logger.Info($"[ZnunySyncFinished] reason={reason} success=true searchFound={searchResultCount} uniqueFound={uniqueTicketCount} created={created} updated={updated} unchanged={unchanged} skipped={skipped}");
            return new ZnunySyncResult(true, true, false, searchResultCount, uniqueTicketCount,
                created, updated, unchanged, skipped, searchLimitReached, string.Empty);
        }
        catch (ZnunySyncBudgetExceededException)
        {
            LastError = "Synchronisierung zum Schutz des Ticketsystems abgebrochen.";
            _logger.Warning($"[ZnunySyncFinished] reason={reason} success=false error=request-budget-exceeded");
            return new ZnunySyncResult(true, false, false, searchResultCount, uniqueTicketCount,
                created, updated, unchanged, skipped, searchLimitReached, LastError);
        }
        catch (ZnunyApiException ex)
        {
            LastError = FormatApiError("Znuny-Synchronisierung fehlgeschlagen", ex);
            LogZnunyError(ex);
            _logger.Warning($"[ZnunySyncFinished] reason={reason} success=false error=api-error");
            return new ZnunySyncResult(true, false, false, searchResultCount, uniqueTicketCount,
                created, updated, unchanged, skipped, searchLimitReached, LastError);
        }
        catch (Exception ex)
        {
            LastError = $"Znuny-Synchronisierung fehlgeschlagen: {ex.Message}";
            _logger.Error($"[ZnunyError] stage=TaskSync errorCode={ex.HResult:X8} message={ex.Message}");
            _logger.Warning($"[ZnunySyncFinished] reason={reason} success=false error=unexpected");
            return new ZnunySyncResult(true, false, false, searchResultCount, uniqueTicketCount,
                created, updated, unchanged, skipped, searchLimitReached, LastError);
        }
        finally
        {
            if (ownsGate) _syncGate.Release();
            if (hasTaskChanges)
                NotifyTasksChanged();
            traffic.Stopwatch.Stop();
            _logger.Info($"[ZnunySyncSummary] reason={traffic.Reason} searchLimit={EffectiveSearchLimit} articleLimit={EffectiveArticleLimit} metadataRemote={traffic.MetadataRequests} detailCacheHits={traffic.DetailCacheHits} detailRemote={traffic.DetailRemote} totalHttpRequests={traffic.TotalRequests} durationMs={traffic.Stopwatch.ElapsedMilliseconds}");
            _logger.Info($"[ZnunySyncTraffic] reason={traffic.Reason} searchRequests={traffic.SearchRequests} ticketMetadataRequests={traffic.MetadataRequests} ticketDetailRequests={traffic.DetailRequests} candidateRequests={traffic.CandidateRequests} totalRequests={traffic.TotalRequests} durationMs={traffic.Stopwatch.ElapsedMilliseconds}");
            _syncTraffic.Value = null;
        }
    }

    private string BuildAssignmentContextKey(int agentId)
    {
        var apiUrl = (_settings.Current.TicketSystemApiUrl ?? string.Empty).Trim().TrimEnd('/').ToLowerInvariant();
        var context = string.Join("|",
            apiUrl,
            agentId.ToString(CultureInfo.InvariantCulture),
            _settings.Current.TicketSystemIncludeOwner ? "owner:1" : "owner:0",
            _settings.Current.TicketSystemIncludeResponsible ? "responsible:1" : "responsible:0");
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(context))).ToLowerInvariant();
    }

    private void CompleteManualSelfAssignmentSuppressions(IReadOnlySet<string> currentAssignedIds)
    {
        foreach (var ticketId in currentAssignedIds)
        {
            if (!_manualSelfAssignmentNotificationSuppressions.TryRemove(ticketId, out _)) continue;
            _logger.Info($"[ZnunySelfAssign] ticketId={ticketId} action=assignment-sync-completed notificationSuppressed=true");
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

    private string ValidateTicketCreateConfiguration(TaskItem task)
    {
        var baseError = ValidateConfiguration(requireAgentId: true);
        if (!string.IsNullOrWhiteSpace(baseError)) return baseError;
        if (string.IsNullOrWhiteSpace(task.Title)) return "Für die Ticket-Erstellung fehlt der Titel der Aufgabe.";
        if (string.IsNullOrWhiteSpace(_settings.Current.TicketSystemTicketCreateRoute)) return "Für die Ticket-Erstellung fehlt die TicketCreate-Route.";
        if (string.IsNullOrWhiteSpace(_settings.Current.TicketSystemTicketCreateMethod)) return "Für die Ticket-Erstellung fehlt die TicketCreate-HTTP-Methode.";
        if (string.IsNullOrWhiteSpace(_settings.Current.TicketSystemCreateQueue))
            return "Für die Ticket-Erstellung fehlt die Standard-Queue. Bitte unter Einstellungen → Ticketsystem konfigurieren.";
        if (string.IsNullOrWhiteSpace(_settings.Current.TicketSystemCreateCustomerUser))
            return "Für die Ticket-Erstellung fehlt der Standard-CustomerUser. Bitte unter Einstellungen → Ticketsystem konfigurieren.";
        return string.Empty;
    }

    private Dictionary<string, object?> BuildTicketCreatePayload(TaskItem task, string sessionId)
    {
        var settings = _settings.Current;
        var ticket = new Dictionary<string, object?>
        {
            ["Title"] = task.Title.Trim(),
            ["Queue"] = settings.TicketSystemCreateQueue,
            ["State"] = settings.TicketSystemCreateState,
            ["Priority"] = settings.TicketSystemCreatePriority,
            ["Lock"] = "unlock",
            ["OwnerID"] = settings.TicketSystemAgentId,
            ["ResponsibleID"] = settings.TicketSystemAgentId,
            ["CustomerUser"] = settings.TicketSystemCreateCustomerUser
        };
        if (!string.IsNullOrWhiteSpace(settings.TicketSystemCreateType))
            ticket["Type"] = settings.TicketSystemCreateType;

        return new Dictionary<string, object?>
        {
            ["SessionID"] = sessionId,
            ["Ticket"] = ticket,
            ["Article"] = new Dictionary<string, object?>
            {
                ["Subject"] = task.Title.Trim(),
                ["Body"] = string.IsNullOrWhiteSpace(task.Description)
                    ? "Ticket aus Plenaro-Aufgabe erstellt."
                    : task.Description,
                ["ContentType"] = "text/plain; charset=utf8",
                ["CommunicationChannel"] = "Internal",
                ["SenderType"] = "agent",
                ["IsVisibleForCustomer"] = 0
            }
        };
    }

    private async Task<string> CreateSessionAsync([CallerMemberName] string caller = "")
    {
        _requestReason.Value = _syncTraffic.Value?.Reason ?? ReasonForCaller(caller);
        if (!string.IsNullOrWhiteSpace(_sessionId) && DateTime.UtcNow < _sessionExpiresUtc)
            return _sessionId;
        await _sessionGate.WaitAsync();
        try
        {
            if (!string.IsNullOrWhiteSpace(_sessionId) && DateTime.UtcNow < _sessionExpiresUtc)
                return _sessionId;
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
        _sessionId = sessionId;
        _sessionExpiresUtc = DateTime.UtcNow.AddMinutes(20);
        return sessionId;
        }
        finally { _sessionGate.Release(); }
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

    private async Task<CandidateDiscoveryResult> RefreshCandidateTicketsCoreAsync(
        string sessionId, string sessionHash, ZnunyRequestReason reason, bool lightweight)
    {
        var stopwatch = Stopwatch.StartNew();
        var includeKeywords = ParseCandidateKeywords(_settings.Current.TicketSystemCandidateKeywords);
        var excludeKeywords = ParseCandidateKeywords(_settings.Current.TicketSystemCandidateExcludeKeywords);
        var candidateUserId = _settings.Current.TicketSystemCandidateUserId;
        _logger.Info($"[ZnunyCandidates] action=refresh-start reason={reason} mode={(lightweight ? "discovery" : "reconcile")} candidateUserId={candidateUserId} includeKeywordCount={includeKeywords.Count} excludeKeywordCount={excludeKeywords.Count}");
        if (includeKeywords.Count == 0)
        {
            _candidateSnapshots.Replace(Array.Empty<ZnunyCandidateTicket>());
            _candidateScanStates.Replace(Array.Empty<CandidateScanState>());
            PublishCandidateTickets(Array.Empty<ZnunyCandidateTicket>(), string.Empty);
            return new CandidateDiscoveryResult(0, 0, 0);
        }

        var ownerIds = await SearchCandidateRoleActiveTicketIdsAsync("Owner", candidateUserId, sessionId, sessionHash);
        var responsibleIds = await SearchCandidateRoleActiveTicketIdsAsync("Responsible", candidateUserId, sessionId, sessionHash);
        var candidateIds = ownerIds.Intersect(responsibleIds, StringComparer.OrdinalIgnoreCase).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var previousState = _candidateScanStates.Load();
        var newIds = candidateIds.Where(id => !previousState.ContainsKey(id)).ToList();
        var idsToEvaluate = lightweight ? newIds : candidateIds.ToList();
        var matches = (lightweight ? _candidateSnapshots.Load().Where(item => candidateIds.Contains(item.TicketId)) : Array.Empty<ZnunyCandidateTicket>()).ToList();
        var matchedById = matches.ToDictionary(item => item.TicketId, StringComparer.OrdinalIgnoreCase);
        var evaluated = new Dictionary<string, (bool Matched, DateTime? Changed)>(StringComparer.OrdinalIgnoreCase);

        foreach (var ticketId in idsToEvaluate.OrderBy(id => id, StringComparer.OrdinalIgnoreCase))
        {
            ZnunyCandidateTicket? candidate = null;
            var cachedEvaluated = false;
            if (lightweight && _detailCache.Load(ticketId) is { Articles.Count: > 0 } cached)
            {
                candidate = EvaluateCachedCandidate(cached, ticketId, includeKeywords, excludeKeywords, candidateUserId);
                cachedEvaluated = true;
            }

            DateTime? remoteChanged = _detailCache.GetRemoteChangedUtc(ticketId);
            if (candidate == null && !cachedEvaluated)
            {
                var ticket = await GetCandidateTicketDetailsAsync(ticketId, sessionId, sessionHash)
                    ?? throw new InvalidOperationException($"TicketGet lieferte keine Daten für TicketID {ticketId}.");
                StoreTicketDetails(ticket);
                remoteChanged = ParseZnunyUtc(ticket.Changed);
                if (ticket.OwnerId == candidateUserId && ticket.ResponsibleId == candidateUserId && !ticket.IsClosed)
                {
                    var match = FindCandidateMatch(ticket, includeKeywords);
                    var exclusion = FindCandidateMatch(ticket, excludeKeywords);
                    if (match.Keyword.Length > 0 && exclusion.Keyword.Length == 0)
                        candidate = CreateCandidate(ticket, candidateUserId, match.Keyword);
                }
            }

            matchedById.Remove(ticketId);
            if (candidate != null) matchedById[ticketId] = candidate;
            evaluated[ticketId] = (candidate != null, remoteChanged);
        }

        var now = DateTime.UtcNow;
        var states = candidateIds.Select(id =>
        {
            if (evaluated.TryGetValue(id, out var item)) return new CandidateScanState(id, now, now, item.Matched, item.Changed);
            var old = previousState[id];
            return old with { LastSeenUtc = now };
        }).ToList();
        var snapshot = matchedById.Values.OrderByDescending(ticket => ticket.TicketNumber, StringComparer.OrdinalIgnoreCase).ToList();
        _candidateScanStates.Replace(states);
        _candidateSnapshots.Replace(snapshot);
        PublishCandidateTickets(snapshot, string.Empty);
        _logger.Info($"[ZnunyCandidates] reason={reason} source={candidateIds.Count} newIds={newIds.Count} evaluated={idsToEvaluate.Count} matched={snapshot.Count} removed={previousState.Keys.Count(id => !candidateIds.Contains(id))} durationMs={stopwatch.ElapsedMilliseconds}");
        return new CandidateDiscoveryResult(newIds.Count, candidateIds.Count - newIds.Count, snapshot.Count);
    }

    private ZnunyCandidateTicket? EvaluateCachedCandidate(TicketBookingContext cached, string ticketId, IReadOnlyList<string> includeKeywords, IReadOnlyList<string> excludeKeywords, int candidateUserId)
    {
        var text = string.Join("\n", new[] { cached.TicketTitle }.Concat(cached.Articles.SelectMany(article => new[] { article.Subject, article.Body })));
        var keyword = includeKeywords.FirstOrDefault(value => text.Contains(value, StringComparison.OrdinalIgnoreCase));
        if (keyword == null || excludeKeywords.Any(value => text.Contains(value, StringComparison.OrdinalIgnoreCase))) return null;
        return new ZnunyCandidateTicket
        {
            TicketId = ticketId, TicketNumber = cached.TicketNumber, Title = cached.TicketTitle,
            DescriptionPreview = CreateDescriptionPreview(cached.Articles.FirstOrDefault()?.Body ?? string.Empty),
            Owner = CandidateDisplayName(candidateUserId), Responsible = CandidateDisplayName(candidateUserId),
            WebUrl = BuildTicketWebUrl(_settings.Current.TicketSystemWebUrl, ticketId), MatchedKeyword = keyword, LastSyncedUtc = DateTime.UtcNow
        };
    }

    private static ZnunyCandidateTicket CreateCandidate(ZnunyTicket ticket, int candidateUserId, string keyword) => new()
    {
        TicketId = ticket.TicketID, TicketNumber = ticket.TicketNumber, Title = ticket.Title,
        DescriptionPreview = CreateDescriptionPreview(ticket.ContentText), Owner = CandidateDisplayName(candidateUserId),
        Responsible = CandidateDisplayName(candidateUserId), State = ticket.State, WebUrl = ticket.WebUrl,
        MatchedKeyword = keyword, LastSyncedUtc = DateTime.UtcNow
    };

    private sealed record CandidateDiscoveryResult(int NewIds, int ExistingIds, int Matched);

    private static string CandidateDisplayName(int candidateUserId)
        => candidateUserId == 1 ? "OTRS, Admin" : $"UserID {candidateUserId}";

    private void PublishCandidateTickets(IReadOnlyList<ZnunyCandidateTicket> tickets, string error)
    {
        _candidateTickets = tickets;
        CandidateTicketsError = error;
        _logger.Info($"[ZnunyCandidatesPublish] serviceCount={tickets.Count}");
        CandidateTicketsChanged?.Invoke();
    }

    private static IReadOnlyList<string> ParseCandidateKeywords(string? value)
        => (value ?? string.Empty)
            .Split(new[] { ',', ';', '\r', '\n', '\t' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

    private static (string Keyword, string Source) FindCandidateMatch(ZnunyTicket ticket, IReadOnlyList<string> keywords)
    {
        foreach (var keyword in keywords)
        {
            if (ticket.Title.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                return (keyword, "Title");
            if (ticket.ArticleSubjects.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                return (keyword, "ArticleSubject");
            if (ticket.ContentText.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                return (keyword, "ArticleBody");
        }

        return (string.Empty, string.Empty);
    }

    private void LogCandidateEvaluation(
        ZnunyTicket ticket,
        string matchedKeyword,
        string matchedIn,
        string result,
        string excludedKeyword = "",
        string excludedIn = "")
    {
        var matchDetails = matchedKeyword.Length == 0
            ? string.Empty
            : $" matchedKeyword='{LogValue(matchedKeyword)}' matchedIn='{matchedIn}'";
        var exclusionDetails = excludedKeyword.Length == 0
            ? string.Empty
            : $" excludedKeyword='{LogValue(excludedKeyword)}' excludedIn='{excludedIn}'";
        _logger.Info($"[ZnunyCandidateEvaluation] ticketId={ticket.TicketID} ticketNumber='{LogValue(ticket.TicketNumber)}' title='{LogValue(ticket.Title)}' owner='{LogValue(ticket.Owner)}' responsible='{LogValue(ticket.Responsible)}' state='{LogValue(ticket.State)}' ownerId={ticket.OwnerId?.ToString(CultureInfo.InvariantCulture) ?? "unknown"} responsibleId={ticket.ResponsibleId?.ToString(CultureInfo.InvariantCulture) ?? "unknown"}{matchDetails}{exclusionDetails} result={result}");
    }

    private static string CreateDescriptionPreview(string text)
    {
        var compact = Regex.Replace(text ?? string.Empty, @"\s+", " ").Trim();
        return compact.Length <= 280 ? compact : compact[..277] + "…";
    }

    private async Task<List<string>> SearchRoleTicketIdsWithOpenCompatibilityAsync(string role, int userId, string sessionId, string sessionHash)
    {
        var onlyOpen = _settings.Current.TicketSystemOnlyOpenTickets && !_settings.Current.TicketSystemShowClosedTickets;
        return await SearchTicketsAsync(role, userId, _settings.Current.TicketSystemTicketSearchRoute,
            _settings.Current.TicketSystemTicketSearchMethod, sessionId, sessionHash, onlyOpen);
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
        LogSearchLimit(isOwner ? "Owner" : "Responsible", ticketIds.Count);
        _logger.Info($"{logTag} method={method} route={route} userId={userId} onlyOpen={onlyOpen} status={(int)result.StatusCode} ticketCount={ticketIds.Count}");
        return ticketIds;
    }

    private async Task<List<string>> SearchCandidateRoleActiveTicketIdsAsync(string role, int userId, string sessionId, string sessionHash)
    {
        var payload = BuildSearchAuthenticationPayload(sessionId);
        ZnunySyncPolicy.ApplyTicketRoleCriteria(payload, role, userId, onlyOpen: true);
        var route = NormalizeRouteValue(_settings.Current.TicketSystemTicketSearchRoute, "/Ticket");
        var stage = $"TicketSearchCandidate{role}Active";
        using var request = BuildSearchRequest("POST", route, payload);
        var result = await SendZnunyAsync(request, stage, "[ZnunyCandidateActiveSearchResponse]", logBody: false);
        var ids = ExtractTicketIdsStrict(result.Body, stage).ToList();
        LogSearchLimit($"Candidate{role}", ids.Count);
        return ids;
    }

    private Dictionary<string, object?> BuildSearchAuthenticationPayload(string sessionId)
    {
        if (string.Equals(_settings.Current.TicketSystemTicketSearchAuthMode, "Direct", StringComparison.OrdinalIgnoreCase))
        {
            var directPayload = new Dictionary<string, object?>
            {
                ["UserLogin"] = _settings.Current.TicketSystemUsername,
                ["Password"] = _settings.GetTicketSystemPassword()
            };
            ZnunySyncPolicy.ApplyTicketSearchLimit(directPayload, _settings.Current.TicketSystemSearchLimit);
            return directPayload;
        }

        var payload = new Dictionary<string, object?> { ["SessionID"] = sessionId };
        ZnunySyncPolicy.ApplyTicketSearchLimit(payload, _settings.Current.TicketSystemSearchLimit);
        return payload;
    }

    private void LogSearchLimit(string operation, int resultCount)
    {
        _logger.Info($"[ZnunyTicketSearch] operation={operation} configuredLimit={_settings.Current.TicketSystemSearchLimit} effectiveLimit={EffectiveSearchLimit} results={resultCount}");
        if (resultCount < EffectiveSearchLimit) return;
        _logger.Warning($"[ZnunySearchLimit] operation={operation} limit={EffectiveSearchLimit} resultCount={resultCount} possiblyTruncated=true sortBy={ZnunySyncPolicy.SearchSortBy} orderBy={ZnunySyncPolicy.SearchOrderBy}");
    }

    private int EffectiveSearchLimit => ZnunySyncPolicy.NormalizeSearchLimit(_settings.Current.TicketSystemSearchLimit);
    private int EffectiveArticleLimit => ZnunySyncPolicy.NormalizeArticleLimit(_settings.Current.TicketSystemArticleLimit);

    private Task<ZnunyTicket?> GetTicketMetadataAsync(string ticketId, string sessionId, string sessionHash)
        => GetTicketAsync(ticketId, sessionId, sessionHash, allArticles: false, dynamicFields: false, "TicketGetMetadata");

    private Task<ZnunyTicket?> GetTicketDetailsAsync(string ticketId, string sessionId, string sessionHash)
        => GetTicketAsync(ticketId, sessionId, sessionHash, allArticles: true, dynamicFields: true, "TicketGetDetails");

    private Task<ZnunyTicket?> GetCandidateTicketDetailsAsync(string ticketId, string sessionId, string sessionHash)
        => GetTicketAsync(ticketId, sessionId, sessionHash, allArticles: true, dynamicFields: false, "CandidateTicketGetDetails");

    private Task<ZnunyTicket?> GetTicketReconciliationDetailsAsync(string ticketId, string sessionId, string sessionHash)
        => GetTicketAsync(ticketId, sessionId, sessionHash, allArticles: true, dynamicFields: true,
            "TicketGetReconciliation", Math.Max(EffectiveArticleLimit, 50));

    private void StoreTicketDetails(ZnunyTicket ticket)
    {
        var articles = ticket.ToArticleItems();
        var (replySource, replyRecipient) = ResolveReplyRecipient(articles, ticket.CustomerUser);
        var context = new TicketBookingContext(ticket.TicketID, ticket.TicketNumber,
            ticket.GetDynamicFieldValue(_settings.Current.TicketSystemCostCenterFieldName),
            ticket.GetDynamicFieldValue(_settings.Current.TicketSystemOrderFieldName),
            Array.Empty<TicketFieldOption>(), Array.Empty<TicketFieldOption>(), string.Empty,
            articles, replySource, replyRecipient, ticket.Title);
        _detailCache.Store(context, ticket.State, ParseZnunyUtc(ticket.Changed));
    }

    private async Task<ZnunyTicket?> GetTicketAsync(string ticketId, string sessionId, string sessionHash, bool allArticles, bool dynamicFields, string operation, int? articleLimit = null)
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
        foreach (var option in ZnunySyncPolicy.TicketGetOptions(allArticles, dynamicFields, articleLimit ?? EffectiveArticleLimit))
            query[option.Key] = option.Value;
        var url = Combine(_settings.Current.TicketSystemApiUrl, route) + ToQueryString(query);
        using var request = new HttpRequestMessage(HttpMethod.Get, url);

        _logger.Info($"[ZnunyTicket] method=GET route={route} ticketId={ticketId} auth={_settings.Current.TicketSystemTicketGetAuthMode} requestedDynamicFields={dynamicFields.ToString().ToLowerInvariant()} allArticles={allArticles.ToString().ToLowerInvariant()} sessionHash={sessionHash}");
        var result = await SendZnunyAsync(request, operation, "[ZnunyTicketResponse]", logBody: false);
        using var doc = JsonDocument.Parse(result.Body);
        ThrowIfApiError(doc.RootElement, "TicketGet");
        var ticketElement = FindFirstTicketElement(doc.RootElement);
        if (!ticketElement.HasValue)
            throw new ZnunyApiException("TicketGet", result.StatusCode, "Protocol", "TicketGet response contains no Ticket object.", result.Body);

        var ticket = ZnunyTicket.FromJson(ticketElement.Value, _settings.Current.TicketSystemWebUrl, allArticles, doc.RootElement);
        _logger.Info($"[ZnunyFirstArticle] ticketId={ticket.TicketID} articleCount={ticket.ArticleCount} selectedArticleId='{ticket.FirstArticleId}' senderType='{ticket.FirstArticleSenderType}' created='{ticket.FirstArticleCreated}' bodyLength={ticket.FirstArticleBody.Length}");
        return ticket;
    }

    private async Task<ZnunyHttpResult> SendZnunyAsync(HttpRequestMessage request, string stage, string responseLogTag, bool logBody = true)
    {
        if (_syncTraffic.Value is { } traffic && !traffic.TryRecord(stage))
        {
            _logger.Warning($"[ZnunyTrafficSafety] action=sync-aborted reason=request-budget-exceeded requests={traffic.TotalRequests}");
            throw new ZnunySyncBudgetExceededException();
        }
        var requestReason = _syncTraffic.Value?.Reason ?? _requestReason.Value ?? InferManualReason(stage);
        var automatic = requestReason.IsAutomatic();
        var ticketId = TicketIdForLog(request.RequestUri);
        _logger.Info($"[ZnunyRequest] reason={requestReason} operation={stage} ticketId={ticketId} automatic={automatic.ToString().ToLowerInvariant()} trafficClass=GenericInterface");
        RecordGlobalTraffic(requestReason, stage);
        using var response = await _client.SendAsync(request);
        var body = await response.Content.ReadAsStringAsync();
        var contentType = response.Content.Headers.ContentType?.ToString() ?? string.Empty;
        _logger.Info(logBody
            ? $"{responseLogTag} status={(int)response.StatusCode} contentType='{contentType}' body={Truncate(RedactSecrets(body))}"
            : $"{responseLogTag} status={(int)response.StatusCode} contentType='{contentType}' bodyLength={body.Length}");

        if (!response.IsSuccessStatusCode)
        {
            var (errorCode, errorMessage) = ExtractApiError(body);
            if (IsInvalidSession(errorCode, errorMessage)) InvalidateSession();
            throw new ZnunyApiException(stage, response.StatusCode, errorCode, string.IsNullOrWhiteSpace(errorMessage) ? response.ReasonPhrase ?? "HTTP error" : errorMessage, body);
        }

        if (TryParseJson(body, out var doc))
        {
            using (doc)
            {
                try { ThrowIfApiError(doc.RootElement, stage, response.StatusCode, body); }
                catch (ZnunyApiException ex)
                {
                    if (IsInvalidSession(ex.ErrorCode, ex.ErrorMessage)) InvalidateSession();
                    throw;
                }
            }
        }

        return new ZnunyHttpResult(response.StatusCode, contentType, body);
    }

    private static ZnunyRequestReason InferManualReason(string operation)
        => operation switch
        {
            "TicketUpdateSelfAssignment" => ZnunyRequestReason.ManualAssign,
            "TicketCreate" => ZnunyRequestReason.ManualTicketCreate,
            "TicketUpdateReply" => ZnunyRequestReason.ManualReply,
            "TicketUpdateTimeBooking" or "TicketUpdateTimeBookingRetry" => ZnunyRequestReason.ManualTimeBooking,
            "SessionGetDiagnostic" => ZnunyRequestReason.ManualConnectionTest,
            _ => ZnunyRequestReason.ManualConnectionTest
        };

    private static ZnunyRequestReason ReasonForCaller(string caller)
        => caller switch
        {
            nameof(AssignCandidateToCurrentAgentAsync) => ZnunyRequestReason.ManualAssign,
            nameof(CreateTicketFromLocalTaskAsync) => ZnunyRequestReason.ManualTicketCreate,
            nameof(RefreshTicketBookingContextAsync) => ZnunyRequestReason.ManualTicketRefresh,
            nameof(RefreshDynamicFieldOptionsAsync) => ZnunyRequestReason.ManualDynamicFieldRefresh,
            nameof(SendTicketReplyAsync) => ZnunyRequestReason.ManualReply,
            nameof(BookTimeAsync) or nameof(RetryTicketTimeBookingAsync) => ZnunyRequestReason.ManualTimeBooking,
            nameof(CheckTicketTimeBookingAsync) => ZnunyRequestReason.ManualBookingCheck,
            nameof(TestRoutesAsync) => ZnunyRequestReason.ManualRouteTest,
            _ => ZnunyRequestReason.ManualConnectionTest
        };

    private static bool IsInvalidSession(string code, string message)
        => code.Contains("Session", StringComparison.OrdinalIgnoreCase)
           && (message.Contains("invalid", StringComparison.OrdinalIgnoreCase)
               || message.Contains("expired", StringComparison.OrdinalIgnoreCase)
               || message.Contains("ungültig", StringComparison.OrdinalIgnoreCase));

    private void InvalidateSession()
    {
        _sessionId = string.Empty;
        _sessionExpiresUtc = DateTime.MinValue;
        _logger.Info("[ZnunySession] action=invalidated retry=false");
    }

    private void LogTraffic(SyncTraffic traffic)
    {
        traffic.Stopwatch.Stop();
        _logger.Info($"[ZnunyActionTraffic] reason={traffic.Reason} requests={traffic.TotalRequests} sessions={traffic.SessionRequests} searches={traffic.SearchRequests} metadataGets={traffic.MetadataRequests} detailGets={traffic.DetailRequests} candidateDetailGets={traffic.CandidateDetailGets} dynamicFieldGets={traffic.DynamicFieldRequests}");
    }

    private void RecordGlobalTraffic(ZnunyRequestReason reason, string operation)
    {
        var now = DateTime.UtcNow;
        _trafficEvents.Enqueue(new ZnunyRequestEvent(now, reason, operation));
        foreach (var minutes in new[] { 1, 5, 15, 0 })
        {
            var events = _trafficEvents.Where(item => minutes == 0 || item.TimestampUtc >= now.AddMinutes(-minutes)).ToList();
            _logger.Info($"[ZnunyTrafficSummary] windowMinutes={(minutes == 0 ? "startup" : minutes.ToString(CultureInfo.InvariantCulture))} total={events.Count} automatic={events.Count(item => item.Reason.IsAutomatic())} manual={events.Count(item => !item.Reason.IsAutomatic())} sessions={events.Count(item => item.Operation == "SessionCreate")} searches={events.Count(item => item.Operation.Contains("Search", StringComparison.OrdinalIgnoreCase))} metadata={events.Count(item => item.Operation == "TicketGetMetadata")} details={events.Count(item => item.Operation.Contains("TicketGetDetails", StringComparison.OrdinalIgnoreCase))} candidates={events.Count(item => item.Operation.Contains("Candidate", StringComparison.OrdinalIgnoreCase))} dynamicFields={events.Count(item => item.Operation == "DynamicFieldOptions")} writes={events.Count(item => item.Operation.Contains("Update", StringComparison.OrdinalIgnoreCase) || item.Operation == "TicketCreate")}");
        }
    }

    private static string TicketIdForLog(Uri? uri)
    {
        if (uri == null) return "-";
        var match = Regex.Match(uri.AbsolutePath, @"/Ticket/(?<id>[^/]+)", RegexOptions.IgnoreCase);
        return match.Success ? Uri.UnescapeDataString(match.Groups["id"].Value) : "-";
    }

    private HttpRequestMessage BuildSearchRequest(string method, string route, Dictionary<string, object?> payload)
    {
        // Central safety net: no TicketSearch can leave this service without a bounded limit.
        ZnunySyncPolicy.ApplyTicketSearchLimit(payload, _settings.Current.TicketSystemSearchLimit);
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
        var payload = new Dictionary<string, object?>();
        ZnunySyncPolicy.ApplyTicketRoleCriteria(payload, owner ? "Owner" : "Responsible", userId, onlyOpen);
        ZnunySyncPolicy.ApplyTicketSearchLimit(payload, _settings.Current.TicketSystemSearchLimit);

        if (string.Equals(_settings.Current.TicketSystemTicketSearchAuthMode, "Direct", StringComparison.OrdinalIgnoreCase))
        {
            payload["UserLogin"] = _settings.Current.TicketSystemUsername;
            payload["Password"] = _settings.GetTicketSystemPassword();
        }
        else
        {
            payload["SessionID"] = sessionId;
        }

        return payload;
    }

    private bool MapTicketToTask(ZnunyTicket ticket, TaskItem task)
    {
        var preserveLocalContent = (task.Tags ?? string.Empty).Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Contains("PlenaroLocalOrigin", StringComparer.OrdinalIgnoreCase);
        var preserveCachedDescription = task.IsZnunyTask && !ticket.HasArticleDetails;
        var title = preserveLocalContent ? task.Title : $"[{ticket.TicketNumber}] {ticket.Title}".Trim();
        var description = preserveLocalContent || preserveCachedDescription ? task.Description : ticket.ToDescription();
        var tags = preserveLocalContent
            ? AddZnunyTicketTags(task.Tags, ticket.TicketID, ticket.TicketNumber)
            : $"Znuny;ZnunyTicketID:{ticket.TicketID};ZnunyTicketNumber:{ticket.TicketNumber}";
        var ticketCreatedUtc = ParseZnunyUtc(ticket.Created);
        var ticketChangedUtc = ParseZnunyUtc(ticket.Changed);
        var changed = !string.Equals(task.Title, title, StringComparison.Ordinal)
                      || !string.Equals(task.Description, description, StringComparison.Ordinal)
                      || !string.Equals(task.TicketUrl, ticket.WebUrl, StringComparison.Ordinal)
                      || !string.Equals(task.Tags, tags, StringComparison.Ordinal)
                      || task.TicketCreatedUtc != ticketCreatedUtc
                      || task.TicketChangedUtc != ticketChangedUtc;
        changed = changed || !string.Equals(task.TicketState, ticket.State, StringComparison.Ordinal)
                          || !string.Equals(task.TicketStateType, ticket.StateType, StringComparison.OrdinalIgnoreCase);

        task.Title = title;
        task.Description = description;
        task.TicketUrl = ticket.WebUrl;
        task.Tags = tags;
        task.TicketCreatedUtc = ticketCreatedUtc;
        task.TicketChangedUtc = ticketChangedUtc;
        task.TicketState = ticket.State;
        task.TicketStateType = ticket.StateType;
        return changed;
    }

    private DateTime? ParseZnunyUtc(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;

        var formats = new[]
        {
            "yyyy-MM-dd HH:mm:ss", "yyyy-MM-dd'T'HH:mm:ss", "yyyy-MM-dd'T'HH:mm:ss.FFFFFFFK",
            "yyyy-MM-dd HH:mm:ssK", "dd.MM.yyyy HH:mm:ss", "dd.MM.yyyy HH:mm"
        };
        if (!DateTimeOffset.TryParseExact(value.Trim(), formats, CultureInfo.InvariantCulture,
                DateTimeStyles.AllowWhiteSpaces, out var withOffset))
        {
            if (!DateTime.TryParseExact(value.Trim(), formats, CultureInfo.InvariantCulture,
                    DateTimeStyles.AllowWhiteSpaces, out var withoutOffset))
                return null;

            withoutOffset = DateTime.SpecifyKind(withoutOffset, DateTimeKind.Unspecified);
            try
            {
                var zone = TimeZoneInfo.FindSystemTimeZoneById(_settings.Current.CalendarTimeZoneId);
                return TimeZoneInfo.ConvertTimeToUtc(withoutOffset, zone);
            }
            catch (TimeZoneNotFoundException)
            {
                _logger.Warning($"[ZnunyDate] timezone='{_settings.Current.CalendarTimeZoneId}' not found; UTC fallback used.");
                return DateTime.SpecifyKind(withoutOffset, DateTimeKind.Utc);
            }
            catch (InvalidTimeZoneException)
            {
                _logger.Warning($"[ZnunyDate] timezone='{_settings.Current.CalendarTimeZoneId}' invalid; UTC fallback used.");
                return DateTime.SpecifyKind(withoutOffset, DateTimeKind.Utc);
            }
        }

        // Exact values without an offset are interpreted in CalendarTimeZoneId, not in the machine culture/time zone.
        if (!Regex.IsMatch(value, @"(?:Z|[+-]\d{2}:?\d{2})\s*$", RegexOptions.IgnoreCase))
        {
            var unspecified = DateTime.SpecifyKind(withOffset.DateTime, DateTimeKind.Unspecified);
            try
            {
                return TimeZoneInfo.ConvertTimeToUtc(unspecified,
                    TimeZoneInfo.FindSystemTimeZoneById(_settings.Current.CalendarTimeZoneId));
            }
            catch (TimeZoneNotFoundException) { return DateTime.SpecifyKind(unspecified, DateTimeKind.Utc); }
            catch (InvalidTimeZoneException) { return DateTime.SpecifyKind(unspecified, DateTimeKind.Utc); }
        }
        return withOffset.UtcDateTime;
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
        var bodyParts = new List<string> { booking.ShortDescription };
        if (!string.IsNullOrWhiteSpace(booking.Note))
            bodyParts.Add($"Notiz:\n{booking.Note}");
        bodyParts.Add(BookingMarker(booking.BookingId));
        var body = string.Join("\n\n", bodyParts);
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

    private static string NormalizeBookingNote(string? note)
    {
        var normalized = note?.Trim() ?? string.Empty;
        return normalized.Length <= 2000 ? normalized : normalized[..2000];
    }

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
        _detailCache.ReplaceFieldOptions(fieldName, DynamicFieldConfigurationFingerprint(), options);
        _logger.Info($"[ZnunyDynamicFieldOptions] field='{fieldName}' optionCount={options.Count} source=Znuny");
        return options;
    }

    private void LoadPersistedDynamicFieldOptions()
    {
        var fields = new[] { _settings.Current.TicketSystemCostCenterFieldName, _settings.Current.TicketSystemOrderFieldName }
            .Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase);
        var loaded = fields.ToDictionary(field => field,
            field => (IReadOnlyList<TicketFieldOption>)_detailCache.LoadFieldOptions(field, DynamicFieldConfigurationFingerprint(), TimeSpan.FromHours(24)),
            StringComparer.OrdinalIgnoreCase);
        if (loaded.Values.Any(options => options.Count > 0))
        {
            _dynamicFieldOptionsCache = loaded;
            _dynamicFieldOptionsCacheExpiresUtc = DateTime.UtcNow.AddHours(24);
            _dynamicFieldOptionsCacheValid = true;
        }
    }

    private string DynamicFieldConfigurationFingerprint()
    {
        var value = $"{_settings.Current.TicketSystemApiUrl}|{_settings.Current.TicketSystemDynamicFieldOptionsRoute}|{_settings.Current.TicketSystemCostCenterFieldName}|{_settings.Current.TicketSystemOrderFieldName}";
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();
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
            string[] values => string.Join(",", values),
            IEnumerable<string> values => string.Join(",", values),
            _ => value.ToString() ?? string.Empty
        };

    private static string ToQueryString(Dictionary<string, string> query)
        => "?" + string.Join("&", query.Select(kvp => $"{Uri.EscapeDataString(kvp.Key)}={Uri.EscapeDataString(kvp.Value)}"));

    private string FormatSearchPayloadForLog(bool owner, int userId, bool onlyOpen, string authMode, string userLogin)
    {
        var idName = owner ? "OwnerIDs" : "ResponsibleIDs";
        var state = onlyOpen ? ",StateType:['new','open']" : string.Empty;
        var auth = string.Equals(authMode, "Direct", StringComparison.OrdinalIgnoreCase)
            ? $"UserLogin:'{userLogin}',Password:'***'"
            : "SessionID:'***'";
        return $"{{{auth},{idName}:[{userId}]{state},Limit:{EffectiveSearchLimit},SortBy:'{ZnunySyncPolicy.SearchSortBy}',OrderBy:'{ZnunySyncPolicy.SearchOrderBy}'}}";
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

    private static string FindStringRecursive(JsonElement element, string name)
    {
        if (element.ValueKind == JsonValueKind.Object)
        {
            foreach (var property in element.EnumerateObject())
            {
                if (string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase)
                    && property.Value.ValueKind is JsonValueKind.String or JsonValueKind.Number)
                    return property.Value.ToString();
                var nested = FindStringRecursive(property.Value, name);
                if (!string.IsNullOrWhiteSpace(nested)) return nested;
            }
        }
        else if (element.ValueKind == JsonValueKind.Array)
        {
            foreach (var item in element.EnumerateArray())
            {
                var nested = FindStringRecursive(item, name);
                if (!string.IsNullOrWhiteSpace(nested)) return nested;
            }
        }
        return string.Empty;
    }

    private static string AddZnunyTicketTags(string? existingTags, string ticketId, string ticketNumber)
    {
        var safeTags = existingTags ?? string.Empty;
        var tags = safeTags
            .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(tag => !tag.StartsWith("ZnunyTicketID:", StringComparison.OrdinalIgnoreCase)
                          && !tag.StartsWith("ZnunyTicketNumber:", StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (!tags.Contains("Znuny", StringComparer.OrdinalIgnoreCase)) tags.Add("Znuny");
        if (!tags.Contains("PlenaroLocalOrigin", StringComparer.OrdinalIgnoreCase)) tags.Add("PlenaroLocalOrigin");
        tags.Add($"ZnunyTicketID:{ticketId}");
        tags.Add($"ZnunyTicketNumber:{ticketNumber}");
        return string.Join(';', tags);
    }

    private static string BuildTicketWebUrl(string webBaseUrl, string ticketId)
    {
        if (string.IsNullOrWhiteSpace(webBaseUrl) || string.IsNullOrWhiteSpace(ticketId)) return webBaseUrl;
        var normalizedWebBaseUrl = RemoveOtrsPathSegmentForTicketUrl(webBaseUrl.Trim());
        var separator = normalizedWebBaseUrl.Contains('?') ? '&' : '?';
        return $"{normalizedWebBaseUrl.TrimEnd('/')}{separator}Action=AgentTicketZoom;TicketID={Uri.EscapeDataString(ticketId)}";
    }

    private static string RemoveOtrsPathSegmentForTicketUrl(string webBaseUrl)
    {
        if (!Uri.TryCreate(webBaseUrl, UriKind.Absolute, out var uri))
            return Regex.Replace(webBaseUrl, "/otrs(?=/|$)", string.Empty, RegexOptions.IgnoreCase);
        var segments = uri.AbsolutePath.Split('/', StringSplitOptions.RemoveEmptyEntries)
            .Where(segment => !string.Equals(segment, "otrs", StringComparison.OrdinalIgnoreCase)).ToArray();
        var path = segments.Length == 0 ? "/" : "/" + string.Join("/", segments);
        var builder = new UriBuilder(uri) { Path = path };
        return builder.Uri.ToString();
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

    private sealed class SyncTraffic
    {
        public ZnunyRequestReason Reason { get; }
        public Stopwatch Stopwatch { get; }
        public int SearchRequests { get; private set; }
        public int MetadataRequests { get; private set; }
        public int DetailRequests { get; private set; }
        public int CandidateRequests { get; private set; }
        public int CandidateDetailGets { get; private set; }
        public int TotalRequests { get; private set; }
        public int DetailCacheHits { get; set; }
        public int DetailRemote { get; set; }
        public int SessionRequests { get; private set; }
        public int DynamicFieldRequests { get; private set; }

        public SyncTraffic(ZnunyRequestReason reason, Stopwatch stopwatch)
        {
            Reason = reason;
            Stopwatch = stopwatch;
        }

        public bool TryRecord(string operation)
        {
            if (Reason.IsAutomatic() && TotalRequests >= ZnunySyncPolicy.MaximumAutomaticRequestsPerSync)
                return false;
            TotalRequests++;
            if (operation == "SessionCreate") SessionRequests++;
            else if (operation == "DynamicFieldOptions") DynamicFieldRequests++;
            else if (operation.Contains("Search", StringComparison.OrdinalIgnoreCase)) { SearchRequests++; if (operation.Contains("Candidate", StringComparison.OrdinalIgnoreCase)) CandidateRequests++; }
            else if (operation.Contains("CandidateTicketGetDetails", StringComparison.OrdinalIgnoreCase)) { CandidateRequests++; CandidateDetailGets++; DetailRequests++; }
            else if (operation == "TicketGetMetadata") MetadataRequests++;
            else if (operation.Contains("TicketGet", StringComparison.OrdinalIgnoreCase)) DetailRequests++;
            return true;
        }
    }

    private sealed class ZnunySyncBudgetExceededException : Exception
    {
        public ZnunySyncBudgetExceededException()
            : base("Der automatische Znuny-Sync wurde wegen eines unerwartet hohen Anfragevolumens abgebrochen.") { }
    }

    private sealed record ZnunyRequestEvent(DateTime TimestampUtc, ZnunyRequestReason Reason, string Operation);

    public void Dispose()
    {
        _initialSyncTimer.Dispose();
        _fullSyncTimer.Dispose();
        _candidateSyncTimer.Dispose();
        _syncGate.Dispose();
        _dynamicFieldOptionsGate.Dispose();
        _sessionGate.Dispose();
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
        public int? StateId { get; init; }
        public string Priority { get; init; } = string.Empty;
        public string Owner { get; init; } = string.Empty;
        public string Responsible { get; init; } = string.Empty;
        public int? OwnerId { get; init; }
        public int? ResponsibleId { get; init; }
        public string Created { get; init; } = string.Empty;
        public string Changed { get; init; } = string.Empty;
        public string DueTime { get; init; } = string.Empty;
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
        public string ArticleSubjects { get; init; } = string.Empty;
        public string ContentText { get; init; } = string.Empty;
        public string CandidateSearchText { get; init; } = string.Empty;
        public string FirstArticleId { get; init; } = string.Empty;
        public string FirstArticleSenderType { get; init; } = string.Empty;
        public string FirstArticleCreated { get; init; } = string.Empty;
        public int ArticleCount { get; init; }
        public bool HasArticleDetails { get; init; }
        public bool IsClosed => IsClosedValue(StateType) || IsClosedValue(State);

        public string GetDynamicFieldValue(string name)
            => string.IsNullOrWhiteSpace(name) || !DynamicFieldValues.TryGetValue(name, out var value) ? string.Empty : value;

        public string FindArticleIdContaining(string marker)
            => Articles.FirstOrDefault(article => article.Body.Contains(marker, StringComparison.OrdinalIgnoreCase))?.ArticleId ?? string.Empty;

        public IReadOnlyList<TicketArticleItem> ToArticleItems()
        {
            var relevant = Articles
                .Where(article => !article.IsSystemArticle && !string.IsNullOrWhiteSpace(article.Body))
                .OrderBy(article => article.CreatedSort)
                .ThenBy(article => article.ArticleId, StringComparer.OrdinalIgnoreCase)
                .ToList();
            return relevant.Select((article, index) =>
            {
                var typeText = GetArticleTypeText(article);
                var subject = string.IsNullOrWhiteSpace(article.Subject) ? "Ohne Betreff" : article.Subject.Trim();
                if (subject.Length > 80) subject = subject[..77] + "…";
                var created = article.CreatedLocal?.ToString("dd.MM.yyyy HH:mm") ?? "Zeit unbekannt";
                return new TicketArticleItem
                {
                    ArticleId = article.ArticleId,
                    CreatedLocal = article.CreatedLocal,
                    Subject = article.Subject,
                    Body = article.Body,
                    SenderType = article.SenderType,
                    CommunicationChannel = article.Channel,
                    From = article.From,
                    To = article.To,
                    ReplyTo = article.ReplyTo,
                    MessageId = article.MessageId,
                    IsVisibleForCustomer = article.IsVisibleForCustomer,
                    TypeText = typeText,
                    DisplayText = $"{index + 1} · {created} · {typeText} · {subject}"
                };
            }).ToList();
        }

        private static string GetArticleTypeText(ZnunyArticle article)
        {
            var customer = article.SenderType.Contains("customer", StringComparison.OrdinalIgnoreCase);
            var agent = article.SenderType.Contains("agent", StringComparison.OrdinalIgnoreCase);
            var email = article.Channel.Contains("email", StringComparison.OrdinalIgnoreCase);
            var internalNote = article.Channel.Contains("internal", StringComparison.OrdinalIgnoreCase)
                               || article.Channel.Contains("note", StringComparison.OrdinalIgnoreCase);
            if (agent && internalNote) return "Interne Notiz";
            if (customer && email) return "Kunde · E-Mail";
            if (customer) return "Kunde";
            if (agent && email) return "Agent · E-Mail";
            if (agent) return "Agent";
            return string.IsNullOrWhiteSpace(article.Channel) ? "Nachricht" : article.Channel;
        }

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

        public static ZnunyTicket FromJson(JsonElement item, string webBaseUrl, bool hasArticleDetails, JsonElement? responseRoot = null)
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
            var articleSubjects = string.Join("\n", articles
                .Where(article => !article.IsSystemArticle && !string.IsNullOrWhiteSpace(article.Subject))
                .Select(article => article.Subject));
            var articleBodies = string.Join("\n\n", articles
                .Where(article => !article.IsSystemArticle && !string.IsNullOrWhiteSpace(article.Body))
                .Select(article => article.Body));
            var title = FirstString(item, "Title");
            return new ZnunyTicket
            {
                TicketID = id,
                TicketNumber = number,
                Title = title,
                Queue = FirstString(item, "Queue"),
                State = FirstString(item, "State"),
                StateType = FirstString(item, "StateType"),
                StateId = FindInteger(item, "StateID"),
                Priority = FirstString(item, "Priority"),
                Owner = FirstString(item, "Owner"),
                Responsible = FirstString(item, "Responsible"),
                OwnerId = FindInteger(item, "OwnerID"),
                ResponsibleId = FindInteger(item, "ResponsibleID"),
                Created = FirstString(item, "Created", "CreateTime"),
                Changed = FirstString(item, "Changed", "ChangeTime"),
                DueTime = FirstString(item, "DueTime", "EscalationTime"),
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
                HasArticleDetails = hasArticleDetails,
                FirstArticleBody = selectedArticle?.Body ?? string.Empty,
                ArticleSubjects = articleSubjects,
                ContentText = articleBodies,
                CandidateSearchText = string.Join("\n", new[] { title, articleSubjects, articleBodies }.Where(value => !string.IsNullOrWhiteSpace(value))),
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
                        NormalizeArticleText(FirstString(element, "Subject"), string.Empty),
                        NormalizeArticleBody(rawBody, contentType),
                        FirstString(element, "CommunicationChannel", "ArticleType", "ArticleTypeID"),
                        FirstString(element, "From"),
                        FirstString(element, "To"),
                        FirstString(element, "ReplyTo", "Reply-To"),
                        FirstString(element, "MessageID", "MessageId"),
                        ParseBoolean(element, "IsVisibleForCustomer"));
                })
                .ToList();
        }

        private static string NormalizeArticleBody(string body, string contentType)
            => NormalizeArticleText(body, contentType, 5000);

        private static string NormalizeArticleText(string value, string contentType, int maximumLength = 1000)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            var text = value;
            if (contentType.Contains("html", StringComparison.OrdinalIgnoreCase) || Regex.IsMatch(text, "<[^>]+>"))
            {
                text = Regex.Replace(text, "<(br|/p|/div|/li|/tr|/h[1-6])[^>]*>", "\n", RegexOptions.IgnoreCase);
                text = Regex.Replace(text, "<li[^>]*>", "- ", RegexOptions.IgnoreCase);
                text = Regex.Replace(text, "<[^>]+>", string.Empty);
            }

            text = WebUtility.HtmlDecode(text);

            text = text.Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n');
            text = Regex.Replace(text, "[ \t]+\n", "\n");
            text = Regex.Replace(text, "\n{3,}", "\n\n").Trim();
            return text.Length <= maximumLength ? text : text[..maximumLength].TrimEnd() + "\n[…]";
        }

        private static bool ParseBoolean(JsonElement element, string propertyName)
        {
            if (!TryGetPropertyCaseInsensitive(element, propertyName, out var value)) return false;
            return value.ValueKind switch
            {
                JsonValueKind.True => true,
                JsonValueKind.Number => value.TryGetInt32(out var number) && number != 0,
                JsonValueKind.String => value.GetString() is { } text
                                        && (text == "1" || bool.TryParse(text, out var parsed) && parsed),
                _ => false
            };
        }

        private sealed record ZnunyArticle(
            string ArticleId,
            string SenderType,
            string Created,
            string Subject,
            string Body,
            string Channel,
            string From,
            string To,
            string ReplyTo,
            string MessageId,
            bool IsVisibleForCustomer)
        {
            public DateTime CreatedSort => DateTime.TryParse(Created, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var parsed)
                ? parsed
                : DateTime.MaxValue;
            public DateTime? CreatedLocal => CreatedSort == DateTime.MaxValue ? null : CreatedSort;

            public bool IsSystemArticle
                => SenderType.Contains("system", StringComparison.OrdinalIgnoreCase)
                   || Channel.Contains("system", StringComparison.OrdinalIgnoreCase)
                   || Channel.Contains("internal", StringComparison.OrdinalIgnoreCase) && string.IsNullOrWhiteSpace(Body);
        }

        private static string BuildTicketWebUrl(string webBaseUrl, string ticketId)
            => TicketSystemService.BuildTicketWebUrl(webBaseUrl, ticketId);

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
