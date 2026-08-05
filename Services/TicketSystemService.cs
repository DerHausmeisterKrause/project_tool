using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
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
    private string _sessionId = string.Empty;
    private int? _userId;

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
        _sessionId = string.Empty;
        _userId = null;

        var interval = Math.Clamp(_settings.Current.TicketSystemSyncIntervalMinutes, 1, 1440);
        _timer.Change(TimeSpan.FromMinutes(interval), TimeSpan.FromMinutes(interval));
    }

    public Task<(int created, int updated, int skipped)> ImportAssignedOpenTicketsAsync()
        => SyncAssignedTicketsAsync("manual");

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
            if (string.IsNullOrWhiteSpace(_settings.Current.TicketSystemUsername))
                return Fail3("Znuny Benutzername fehlt.");
            if (string.IsNullOrWhiteSpace(_settings.GetTicketSystemPassword()))
                return Fail3("Znuny Passwort fehlt.");
            if (!_settings.Current.TicketSystemIncludeOwner && !_settings.Current.TicketSystemIncludeResponsible)
                return Fail3("Znuny Sync benötigt Owner oder Responsible als Suchkriterium.");

            _logger.Info($"[Znuny] Sync start reason={reason} baseUrl='{SanitizeUrl(_settings.Current.TicketSystemApiUrl)}' onlyOpen={_settings.Current.TicketSystemOnlyOpenTickets} showClosed={_settings.Current.TicketSystemShowClosedTickets} includeOwner={_settings.Current.TicketSystemIncludeOwner} includeResponsible={_settings.Current.TicketSystemIncludeResponsible}");
            await EnsureSessionAsync(forceRenew: false);
            _userId ??= await ResolveUserIdAsync();
            if (!_userId.HasValue)
                return Fail3("Znuny UserID konnte nicht automatisch aus SessionGet ermittelt werden.");

            var ticketIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (_settings.Current.TicketSystemIncludeOwner)
                foreach (var id in await SearchTicketsAsync("Owner", _userId.Value)) ticketIds.Add(id);
            if (_settings.Current.TicketSystemIncludeResponsible)
                foreach (var id in await SearchTicketsAsync("Responsible", _userId.Value)) ticketIds.Add(id);

            var existing = _tasks.GetAllTasks()
                .Where(t => !string.IsNullOrWhiteSpace(ExtractZnunyTicketIdFromTask(t)))
                .GroupBy(ExtractZnunyTicketIdFromTask, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var ticketId in ticketIds)
            {
                var ticket = await GetTicketAsync(ticketId);
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
        catch (Exception ex)
        {
            _sessionId = string.Empty;
            _userId = null;
            LastError = $"Znuny Sync fehlgeschlagen: {ex.Message}";
            _logger.Error($"[ZnunyError] {ex}");
            return (created, updated, skipped);
        }
        finally
        {
            _syncGate.Release();
        }
    }

    private async Task EnsureSessionAsync(bool forceRenew)
    {
        if (!forceRenew && !string.IsNullOrWhiteSpace(_sessionId))
            return;

        _sessionId = string.Empty;
        _userId = null;

        var url = Combine(_settings.Current.TicketSystemApiUrl, "Session");
        var payload = new Dictionary<string, object?>
        {
            ["UserLogin"] = _settings.Current.TicketSystemUsername,
            ["Password"] = _settings.GetTicketSystemPassword()
        };

        _logger.Info($"[ZnunyLogin] POST {SanitizeUrl(url)} payload={{UserLogin:'{_settings.Current.TicketSystemUsername}',Password:'***'}}");
        using var response = await PostJsonAsync(url, payload);
        var json = await response.Content.ReadAsStringAsync();
        _logger.Info($"[ZnunyLogin] status={(int)response.StatusCode} response={Truncate(json)}");
        response.EnsureSuccessStatusCode();

        using var doc = JsonDocument.Parse(json);
        _sessionId = FirstString(doc.RootElement, "SessionID");
        if (string.IsNullOrWhiteSpace(_sessionId))
            throw new InvalidOperationException("Znuny SessionCreate lieferte keine SessionID.");
    }

    private async Task<int?> ResolveUserIdAsync()
    {
        await EnsureSessionAsync(false);
        _logger.Info("[Znuny] Resolve UserID via SessionGet");
        var json = await GetStringWithSessionRetryAsync(() => Combine(_settings.Current.TicketSystemApiUrl, $"Session/SessionID={Uri.EscapeDataString(_sessionId)}"));
        _logger.Info($"[Znuny] SessionGet response={Truncate(json)}");
        using var doc = JsonDocument.Parse(json);
        return FindSessionValue(doc.RootElement, "UserID", "UserId", "UserIDRaw") ?? FindInteger(doc.RootElement, "UserID", "UserId");
    }

    private async Task<IEnumerable<string>> SearchTicketsAsync(string role, int userId)
    {
        await EnsureSessionAsync(false);
        var url = Combine(_settings.Current.TicketSystemApiUrl, "Ticket/Search");
        var payload = new Dictionary<string, object?> { ["SessionID"] = _sessionId };
        payload[role == "Owner" ? "OwnerIDs" : "ResponsibleIDs"] = new[] { userId };
        if (_settings.Current.TicketSystemOnlyOpenTickets && !_settings.Current.TicketSystemShowClosedTickets)
            payload["StateType"] = new[] { "Open" };

        var logTag = role == "Owner" ? "[ZnunySearchOwner]" : "[ZnunySearchResponsible]";
        _logger.Info($"{logTag} POST {SanitizeUrl(url)} payload={{SessionID:'***',{(role == "Owner" ? "OwnerIDs" : "ResponsibleIDs")}:[{userId}],StateType:'{payload.GetValueOrDefault("StateType")}'}}");
        var json = await PostJsonWithSessionRetryAsync(url, payload);
        _logger.Info($"{logTag} response={Truncate(json)}");
        using var doc = JsonDocument.Parse(json);
        return ExtractTicketIds(doc.RootElement);
    }

    private async Task<ZnunyTicket?> GetTicketAsync(string ticketId)
    {
        await EnsureSessionAsync(false);
        _logger.Info($"[ZnunyTicket] GET Ticket/{ticketId}");
        var json = await GetStringWithSessionRetryAsync(() => Combine(_settings.Current.TicketSystemApiUrl, $"Ticket/{Uri.EscapeDataString(ticketId)}?SessionID={Uri.EscapeDataString(_sessionId)}&DynamicFields=1"));
        _logger.Info($"[ZnunyTicket] ticketId={ticketId} response={Truncate(json)}");
        using var doc = JsonDocument.Parse(json);
        var ticketElement = FindFirstTicketElement(doc.RootElement);
        return ticketElement.HasValue ? ZnunyTicket.FromJson(ticketElement.Value, _settings.Current.TicketSystemWebUrl) : null;
    }

    private async Task<string> PostJsonWithSessionRetryAsync(string url, Dictionary<string, object?> payload)
    {
        using var response = await PostJsonAsync(url, payload);
        var json = await response.Content.ReadAsStringAsync();
        if (IsSessionExpired(response, json))
        {
            await EnsureSessionAsync(forceRenew: true);
            payload["SessionID"] = _sessionId;
            using var retry = await PostJsonAsync(url, payload);
            var retryJson = await retry.Content.ReadAsStringAsync();
            retry.EnsureSuccessStatusCode();
            return retryJson;
        }

        response.EnsureSuccessStatusCode();
        return json;
    }

    private async Task<string> GetStringWithSessionRetryAsync(Func<string> urlFactory)
    {
        var url = urlFactory();
        using var response = await _client.GetAsync(url);
        var json = await response.Content.ReadAsStringAsync();
        if (IsSessionExpired(response, json))
        {
            await EnsureSessionAsync(forceRenew: true);
            url = urlFactory();
            using var retry = await _client.GetAsync(url);
            var retryJson = await retry.Content.ReadAsStringAsync();
            retry.EnsureSuccessStatusCode();
            return retryJson;
        }

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

    private static bool IsSessionExpired(HttpResponseMessage response, string json)
        => response.StatusCode is HttpStatusCode.Unauthorized or HttpStatusCode.Forbidden
           || json.Contains("Session", StringComparison.OrdinalIgnoreCase) && (json.Contains("invalid", StringComparison.OrdinalIgnoreCase) || json.Contains("expired", StringComparison.OrdinalIgnoreCase));

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
        foreach (var name in names)
        {
            if (!item.TryGetProperty(name, out var value)) continue;
            if (value.ValueKind == JsonValueKind.Number && value.TryGetInt32(out var number)) return number;
            if (int.TryParse(FirstString(item, name), out var parsed)) return parsed;
        }

        return null;
    }

    private static string FirstString(JsonElement item, params string[] names)
    {
        if (item.ValueKind != JsonValueKind.Object) return string.Empty;
        foreach (var name in names)
        {
            if (!item.TryGetProperty(name, out var value)) continue;
            if (value.ValueKind == JsonValueKind.String) return value.GetString() ?? string.Empty;
            if (value.ValueKind is JsonValueKind.Number or JsonValueKind.True or JsonValueKind.False) return value.ToString();
        }
        return string.Empty;
    }

    private static string Combine(string baseUrl, string relative) => $"{baseUrl.TrimEnd('/')}/{relative.TrimStart('/')}";
    private static string SanitizeUrl(string value) => value.Replace("Password=", "Password=***", StringComparison.OrdinalIgnoreCase).Replace("SessionID=", "SessionID=***", StringComparison.OrdinalIgnoreCase);
    private static string Truncate(string value) => value.Length <= 3000 ? value : value[..3000] + "...";

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
