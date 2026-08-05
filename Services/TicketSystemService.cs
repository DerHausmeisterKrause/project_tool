using System.Net.Http;
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
    private readonly System.Threading.Timer _timer;

    public string LastError { get; private set; } = string.Empty;

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

    public async Task<(bool success, string message)> TestConnectionAsync()
    {
        LastError = string.Empty;

        try
        {
            if (string.IsNullOrWhiteSpace(_settings.Current.TicketSystemApiUrl) || IsPlaceholderUrl(_settings.Current.TicketSystemApiUrl))
                return (false, "Bitte zuerst die echte Znuny API-URL in den Einstellungen eintragen.");
            if (string.IsNullOrWhiteSpace(_settings.Current.TicketSystemUsername))
                return (false, "Znuny Benutzername fehlt.");
            if (string.IsNullOrWhiteSpace(_settings.GetTicketSystemPassword()))
                return (false, "Znuny Passwort fehlt.");

            var sessionId = await CreateSessionAsync();
            var sessionHash = HashSessionId(sessionId);
            var agentId = GetConfiguredAgentId();
            var sessionGetInfo = "SessionGet übersprungen, Agenten-ID ist konfiguriert.";

            if (agentId.HasValue)
            {
                _logger.Info($"[ZnunyUser] source=ConfiguredSettings userId={agentId.Value}");
                var optionalUserId = await TryResolveUserIdFromSessionAsync(sessionId, sessionHash);
                if (optionalUserId.HasValue)
                    sessionGetInfo = $"SessionGet optional erfolgreich, UserID={optionalUserId.Value}.";
            }
            else
            {
                agentId = await TryResolveUserIdFromSessionAsync(sessionId, sessionHash);
                if (!agentId.HasValue)
                    return (false, "Login erfolgreich, aber die Znuny-Agenten-ID konnte nicht automatisch ermittelt werden. Bitte trage sie in den Einstellungen ein.");

                sessionGetInfo = $"SessionGet erfolgreich, UserID={agentId.Value}.";
                _logger.Info($"[ZnunyUser] source=SessionGet userId={agentId.Value}");
            }

            var ownerCount = (await SearchTicketsAsync("Owner", sessionId, sessionHash, agentId.Value)).Count();
            var responsibleCount = (await SearchTicketsAsync("Responsible", sessionId, sessionHash, agentId.Value)).Count();

            return (true, $"Login erfolgreich. Agenten-ID: {agentId.Value}. Owner-Tickets: {ownerCount}. Responsible-Tickets: {responsibleCount}. {sessionGetInfo}");
        }
        catch (ZnunyApiException ex)
        {
            _logger.Error($"[ZnunyError] stage={ex.Stage} errorCode={ex.ErrorCode} message={ex.ErrorMessage}");
            return (false, $"Znuny-Test fehlgeschlagen ({ex.Stage}): {ex.ErrorMessage}");
        }
        catch (Exception ex)
        {
            _logger.Error($"[ZnunyError] stage=Login errorCode={ex.HResult:X8} message={ex.Message}");
            return (false, $"Znuny-Test fehlgeschlagen: {ex.Message}");
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

        try
        {
            if (string.IsNullOrWhiteSpace(_settings.Current.TicketSystemApiUrl))
                return Fail3("Znuny Server URL fehlt.");
            if (IsPlaceholderUrl(_settings.Current.TicketSystemApiUrl))
                return Fail3("Bitte die Znuny API-URL in den Einstellungen anpassen und SERVER durch den echten Hostnamen ersetzen.");
            if (string.IsNullOrWhiteSpace(_settings.Current.TicketSystemUsername))
                return Fail3("Znuny Benutzername fehlt.");
            if (string.IsNullOrWhiteSpace(_settings.GetTicketSystemPassword()))
                return Fail3("Znuny Passwort fehlt.");
            if (!_settings.Current.TicketSystemIncludeOwner && !_settings.Current.TicketSystemIncludeResponsible)
                return Fail3("Znuny Sync benötigt Owner oder Responsible als Suchkriterium.");

            _logger.Info($"[Znuny] Sync start reason={reason} baseUrl='{SanitizeUrl(_settings.Current.TicketSystemApiUrl)}' onlyOpen={_settings.Current.TicketSystemOnlyOpenTickets} showClosed={_settings.Current.TicketSystemShowClosedTickets} includeOwner={_settings.Current.TicketSystemIncludeOwner} includeResponsible={_settings.Current.TicketSystemIncludeResponsible}");
            var sessionId = await CreateSessionAsync();
            var sessionHash = HashSessionId(sessionId);
            var userId = GetConfiguredAgentId();
            if (userId.HasValue)
            {
                _logger.Info($"[ZnunyUser] source=ConfiguredSettings userId={userId.Value}");
            }
            else
            {
                userId = await TryResolveUserIdFromSessionAsync(sessionId, sessionHash);
                if (!userId.HasValue)
                    return Fail3("Die Znuny-Agenten-ID konnte nicht automatisch ermittelt werden. Bitte trage sie in den Einstellungen ein.");

                _logger.Info($"[ZnunyUser] source=SessionGet userId={userId.Value}");
            }

            var ticketIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (_settings.Current.TicketSystemIncludeOwner)
                foreach (var id in await SearchTicketsAsync("Owner", sessionId, sessionHash, userId.Value)) ticketIds.Add(id);
            if (_settings.Current.TicketSystemIncludeResponsible)
                foreach (var id in await SearchTicketsAsync("Responsible", sessionId, sessionHash, userId.Value)) ticketIds.Add(id);

            var existing = _tasks.GetAllTasks()
                .Where(t => !string.IsNullOrWhiteSpace(ExtractZnunyTicketIdFromTask(t)))
                .GroupBy(ExtractZnunyTicketIdFromTask, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var ticketId in ticketIds)
            {
                var ticket = await GetTicketAsync(ticketId, sessionId, sessionHash);
                if (ticket == null)
                {
                    skipped++;
                    continue;
                }

                if (ticket.IsClosed && _settings.Current.TicketSystemOnlyOpenTickets && !_settings.Current.TicketSystemShowClosedTickets)
                {
                    skipped++;
                    continue;
                }

                seen.Add(ticket.TicketID);
                if (existing.TryGetValue(ticket.TicketID, out var task))
                {
                    MapTicketToTask(ticket, task);
                    _tasks.UpdateTask(task);
                    updated++;
                    _logger.Info($"[ZnunyTaskUpdated] ticketId={ticket.TicketID} ticketNumber='{ticket.TicketNumber}' taskId={task.Id}");
                }
                else
                {
                    task = new TaskItem();
                    MapTicketToTask(ticket, task);
                    _tasks.CreateTask(task);
                    created++;
                    _logger.Info($"[ZnunyTaskCreated] ticketId={ticket.TicketID} ticketNumber='{ticket.TicketNumber}' taskId={task.Id}");
                }
            }

            if (_settings.Current.TicketSystemOnlyOpenTickets)
            {
                foreach (var task in existing.Values.Where(t => !seen.Contains(ExtractZnunyTicketIdFromTask(t))))
                {
                    if (task.Status == TaskStatus.Done) continue;
                    task.Status = TaskStatus.Done;
                    _tasks.UpdateTask(task);
                    updated++;
                    _logger.Info($"[ZnunyTaskUpdated] missingOpenTicket taskId={task.Id} action=MarkedDone");
                }
            }

            _logger.Info($"[ZnunySyncFinished] created={created} updated={updated} skipped={skipped} totalTickets={ticketIds.Count}");
            return (created, updated, skipped);
        }
        catch (ZnunyApiException ex)
        {
            LastError = $"Znuny Sync fehlgeschlagen: {ex.Message}";
            _logger.Error($"[ZnunyError] stage={ex.Stage} errorCode={ex.ErrorCode} message={ex.ErrorMessage}");
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
        }
    }

    private async Task<string> CreateSessionAsync()
    {
        var url = Combine(_settings.Current.TicketSystemApiUrl, "Session");
        var payload = new Dictionary<string, object?>
        {
            ["UserLogin"] = _settings.Current.TicketSystemUsername,
            ["Password"] = _settings.GetTicketSystemPassword()
        };

        _logger.Info($"[ZnunyLogin] POST {SanitizeUrl(url)} payload={{UserLogin:'{_settings.Current.TicketSystemUsername}',Password:'***'}}");
        using var response = await PostJsonAsync(url, payload);
        var json = await response.Content.ReadAsStringAsync();
        _logger.Info($"[ZnunyLogin] status={(int)response.StatusCode} response={Truncate(RedactSecrets(json))}");
        response.EnsureSuccessStatusCode();

        using var doc = JsonDocument.Parse(json);
        ThrowIfApiError(doc.RootElement, "Login");
        var sessionId = FirstString(doc.RootElement, "SessionID");
        if (string.IsNullOrWhiteSpace(sessionId))
            throw new InvalidOperationException("Znuny SessionCreate lieferte keine SessionID.");

        _logger.Info($"[ZnunySession] sessionCreated=True sessionHash={HashSessionId(sessionId)}");
        return sessionId;
    }

    private async Task<JsonDocument> GetSessionAsync(string sessionId, string sessionHash)
    {
        var url = Combine(_settings.Current.TicketSystemApiUrl, $"Session/SessionID={Uri.EscapeDataString(sessionId)}");
        _logger.Info($"[ZnunySession] GET {SanitizeUrl(url)} createdSessionIdHash={sessionHash} reusedFor=SessionGet");
        using var response = await _client.GetAsync(url);
        var json = await response.Content.ReadAsStringAsync();
        _logger.Info($"[ZnunySession] SessionGet status={(int)response.StatusCode} response={Truncate(RedactSecrets(json))}");
        response.EnsureSuccessStatusCode();

        var doc = JsonDocument.Parse(json);
        if (ContainsError(doc.RootElement, out var errorCode, out var errorMessage))
        {
            doc.Dispose();
            throw new ZnunyApiException("SessionGet", errorCode, errorMessage);
        }

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
            _logger.Error($"[ZnunyError] stage={ex.Stage} errorCode={ex.ErrorCode} message={ex.ErrorMessage}");
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

    private async Task<IEnumerable<string>> SearchTicketsAsync(string role, string sessionId, string sessionHash, int userId)
    {
        var url = Combine(_settings.Current.TicketSystemApiUrl, "Ticket/Search");
        var payload = new Dictionary<string, object?> { ["SessionID"] = sessionId };
        payload[role == "Owner" ? "OwnerIDs" : "ResponsibleIDs"] = new[] { userId };
        if (_settings.Current.TicketSystemOnlyOpenTickets && !_settings.Current.TicketSystemShowClosedTickets)
            payload["StateType"] = "Open";

        var logTag = role == "Owner" ? "[ZnunySearchOwner]" : "[ZnunySearchResponsible]";
        var stage = role == "Owner" ? "OwnerSearch" : "ResponsibleSearch";
        _logger.Info($"{logTag} method=POST route=/Ticket/Search userId={userId} onlyOpen={_settings.Current.TicketSystemOnlyOpenTickets && !_settings.Current.TicketSystemShowClosedTickets} sessionHash={sessionHash}");
        var json = await PostJsonAsync(url, payload, logTag);
        _logger.Info($"[ZnunySearchResponse] stage={stage} response={Truncate(RedactSecrets(json))}");
        using var doc = JsonDocument.Parse(json);
        ThrowIfApiError(doc.RootElement, stage);
        var ticketIds = ExtractTicketIds(doc.RootElement).ToList();
        _logger.Info($"{logTag} method=POST route=/Ticket/Search userId={userId} onlyOpen={_settings.Current.TicketSystemOnlyOpenTickets && !_settings.Current.TicketSystemShowClosedTickets} status=OK ticketCount={ticketIds.Count}");
        return ticketIds;
    }

    private async Task<ZnunyTicket?> GetTicketAsync(string ticketId, string sessionId, string sessionHash)
    {
        _logger.Info($"[ZnunyTicket] GET Ticket/{ticketId} createdSessionIdHash={sessionHash}");
        var url = Combine(_settings.Current.TicketSystemApiUrl, $"Ticket/{Uri.EscapeDataString(ticketId)}?SessionID={Uri.EscapeDataString(sessionId)}&DynamicFields=1");
        var json = await GetStringAsync(url, "[ZnunyTicket]");
        _logger.Info($"[ZnunyTicket] ticketId={ticketId} response={Truncate(RedactSecrets(json))}");
        using var doc = JsonDocument.Parse(json);
        ThrowIfApiError(doc.RootElement, "TicketGet");
        var ticketElement = FindFirstTicketElement(doc.RootElement);
        return ticketElement.HasValue ? ZnunyTicket.FromJson(ticketElement.Value, _settings.Current.TicketSystemWebUrl) : null;
    }

    private async Task<string> PostJsonAsync(string url, Dictionary<string, object?> payload, string logTag)
    {
        using var response = await PostJsonAsync(url, payload);
        var json = await response.Content.ReadAsStringAsync();
        _logger.Info($"{logTag} status={(int)response.StatusCode}");
        response.EnsureSuccessStatusCode();
        return json;
    }

    private async Task<string> GetStringAsync(string url, string logTag)
    {
        _logger.Info($"{logTag} GET {SanitizeUrl(url)}");
        using var response = await _client.GetAsync(url);
        var json = await response.Content.ReadAsStringAsync();
        _logger.Info($"{logTag} status={(int)response.StatusCode}");
        response.EnsureSuccessStatusCode();
        return json;
    }

    private Task<HttpResponseMessage> PostJsonAsync(string url, object payload)
    {
        var json = JsonSerializer.Serialize(payload);
        return _client.PostAsync(url, new StringContent(json, Encoding.UTF8, "application/json"));
    }

    private void MapTicketToTask(ZnunyTicket ticket, TaskItem task)
    {
        task.Title = $"[{ticket.TicketNumber}] {ticket.Title}".Trim();
        task.Description = ticket.ToDescription();
        task.TicketUrl = ticket.WebUrl;
        task.Status = ticket.IsClosed ? TaskStatus.Done : TaskStatus.Planned;
        task.Tags = $"Znuny;ZnunyTicketID:{ticket.TicketID};ZnunyTicketNumber:{ticket.TicketNumber}";
    }

    private static string ExtractZnunyTicketIdFromTask(TaskItem task)
    {
        var parts = (task.Tags ?? string.Empty).Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var id = parts.FirstOrDefault(p => p.StartsWith("ZnunyTicketID:", StringComparison.OrdinalIgnoreCase));
        return id?.Split(':', 2).ElementAtOrDefault(1) ?? string.Empty;
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

        if (!root.TryGetProperty("Error", out var error) || error.ValueKind != JsonValueKind.Object)
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

    private static void ThrowIfApiError(JsonElement root, string stage)
    {
        if (ContainsError(root, out var errorCode, out var errorMessage))
            throw new ZnunyApiException(stage, errorCode, errorMessage);
    }

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
    private static string RedactSecrets(string value)
        => Regex.Replace(value, "\"SessionID\"\\s*:\\s*\"[^\"]+\"", "\"SessionID\":\"***\"", RegexOptions.IgnoreCase);

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
        _client.Dispose();
    }

    private sealed class ZnunyApiException : Exception
    {
        public string Stage { get; }
        public string ErrorCode { get; }
        public string ErrorMessage { get; }

        public ZnunyApiException(string stage, string errorCode, string errorMessage)
            : base(string.IsNullOrWhiteSpace(errorCode) ? errorMessage : $"{errorCode}: {errorMessage}")
        {
            Stage = stage;
            ErrorCode = errorCode;
            ErrorMessage = errorMessage;
        }
    }

    private sealed class ZnunyTicket
    {
        public string TicketID { get; init; } = string.Empty;
        public string TicketNumber { get; init; } = string.Empty;
        public string Title { get; init; } = string.Empty;
        public string Queue { get; init; } = string.Empty;
        public string State { get; init; } = string.Empty;
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
        public bool IsClosed => State.Contains("closed", StringComparison.OrdinalIgnoreCase) || State.Contains("removed", StringComparison.OrdinalIgnoreCase) || State.Contains("merged", StringComparison.OrdinalIgnoreCase);

        public static ZnunyTicket FromJson(JsonElement item, string webBaseUrl)
        {
            var id = FirstString(item, "TicketID");
            var number = FirstString(item, "TicketNumber");
            return new ZnunyTicket
            {
                TicketID = id,
                TicketNumber = number,
                Title = FirstString(item, "Title"),
                Queue = FirstString(item, "Queue"),
                State = FirstString(item, "State"),
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
                DynamicFields = ExtractDynamicFields(item)
            };
        }

        public string ToDescription()
        {
            var sb = new StringBuilder();
            sb.AppendLine($"Znuny TicketID: {TicketID}");
            sb.AppendLine($"TicketNumber: {TicketNumber}");
            sb.AppendLine($"Title: {Title}");
            sb.AppendLine($"Queue: {Queue}");
            sb.AppendLine($"State: {State}");
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

        private static string BuildTicketWebUrl(string webBaseUrl, string ticketId)
        {
            if (string.IsNullOrWhiteSpace(webBaseUrl) || string.IsNullOrWhiteSpace(ticketId)) return webBaseUrl;
            var separator = webBaseUrl.Contains('?') ? '&' : '?';
            return $"{webBaseUrl.TrimEnd('/')}{separator}Action=AgentTicketZoom;TicketID={Uri.EscapeDataString(ticketId)}";
        }

        private static string ExtractDynamicFields(JsonElement item)
        {
            if (!item.TryGetProperty("DynamicField", out var value)) return string.Empty;
            return value.ToString();
        }
    }
}
