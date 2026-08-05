using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using TaskTool.Models;
using TaskStatus = TaskTool.Models.TaskStatus;

namespace TaskTool.Services;

public class TicketSystemService
{
    private readonly SettingsService _settings;
    private readonly TaskService _tasks;
    private readonly LoggerService _logger;

    public string LastError { get; private set; } = string.Empty;

    public TicketSystemService(SettingsService settings, TaskService tasks, LoggerService logger)
    {
        _settings = settings;
        _tasks = tasks;
        _logger = logger;
    }

    public async Task<(int created, int skipped)> ImportAssignedOpenTicketsAsync()
    {
        LastError = string.Empty;

        if (string.IsNullOrWhiteSpace(_settings.Current.TicketSystemApiUrl))
            return Fail("Ticketsystem API-URL fehlt.");

        try
        {
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(30) };
            ConfigureAuthentication(client);

            using var response = await client.GetAsync(_settings.Current.TicketSystemApiUrl);
            var json = await response.Content.ReadAsStringAsync();
            if (!response.IsSuccessStatusCode)
                return Fail($"Ticketsystem API Fehler {(int)response.StatusCode}: {response.ReasonPhrase}");

            var tickets = ParseTickets(json);
            var existingTasks = _tasks.GetAllTasks();
            var existingUrls = existingTasks
                .Where(t => !string.IsNullOrWhiteSpace(t.TicketUrl))
                .Select(t => t.TicketUrl.Trim())
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            var existingExternalIds = existingTasks
                .SelectMany(t => (t.Tags ?? string.Empty).Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var created = 0;
            var skipped = 0;
            foreach (var ticket in tickets)
            {
                if (string.IsNullOrWhiteSpace(ticket.Title))
                {
                    skipped++;
                    continue;
                }

                var externalId = ticket.ExternalId.Trim();
                if ((!string.IsNullOrWhiteSpace(ticket.Url) && existingUrls.Contains(ticket.Url.Trim()))
                    || (!string.IsNullOrWhiteSpace(externalId) && existingExternalIds.Contains(externalId)))
                {
                    skipped++;
                    continue;
                }

                var task = new TaskItem
                {
                    Title = ticket.Title,
                    Description = ticket.Description,
                    TicketUrl = ticket.Url,
                    Status = TaskStatus.Planned,
                    Tags = string.IsNullOrWhiteSpace(externalId) ? "Ticketsystem" : $"Ticketsystem;{externalId}"
                };

                _tasks.CreateTask(task);
                if (!string.IsNullOrWhiteSpace(task.TicketUrl))
                    existingUrls.Add(task.TicketUrl.Trim());
                if (!string.IsNullOrWhiteSpace(externalId))
                    existingExternalIds.Add(externalId);
                created++;
            }

            _logger.Info($"[TicketSystem] Imported tickets created={created} skipped={skipped} source='{_settings.Current.TicketSystemApiUrl}'");
            return (created, skipped);
        }
        catch (Exception ex)
        {
            _logger.Error($"[TicketSystem] Import failed: {ex}");
            return Fail($"Ticketsystem Import fehlgeschlagen: {ex.Message}");
        }
    }

    private void ConfigureAuthentication(HttpClient client)
    {
        var token = _settings.Current.TicketSystemApiToken?.Trim();
        if (!string.IsNullOrWhiteSpace(token))
        {
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);
            return;
        }

        var username = _settings.Current.TicketSystemUsername?.Trim();
        var password = _settings.Current.TicketSystemPassword ?? string.Empty;
        if (!string.IsNullOrWhiteSpace(username) && !string.IsNullOrWhiteSpace(password))
        {
            var raw = Convert.ToBase64String(Encoding.UTF8.GetBytes($"{username}:{password}"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", raw);
        }
    }

    private static List<TicketSystemTicket> ParseTickets(string json)
    {
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        var source = FindTicketArray(root);
        var tickets = new List<TicketSystemTicket>();

        foreach (var item in source.EnumerateArray())
        {
            var externalId = FirstString(item, "id", "key", "number", "ticketId", "ticket_id");
            var title = FirstString(item, "title", "summary", "subject", "name");
            var description = FirstString(item, "description", "body", "text");
            var url = FirstString(item, "url", "html_url", "web_url", "self");

            tickets.Add(new TicketSystemTicket
            {
                ExternalId = externalId,
                Title = title,
                Description = description,
                Url = url
            });
        }

        return tickets;
    }

    private static JsonElement FindTicketArray(JsonElement root)
    {
        if (root.ValueKind == JsonValueKind.Array)
            return root;

        foreach (var propertyName in new[] { "tickets", "issues", "data", "items", "results" })
        {
            if (root.ValueKind == JsonValueKind.Object
                && root.TryGetProperty(propertyName, out var value)
                && value.ValueKind == JsonValueKind.Array)
            {
                return value;
            }
        }

        throw new InvalidOperationException("API-Antwort enthält keine Ticketliste. Erwartet wird ein Array oder ein Objekt mit tickets/issues/data/items/results.");
    }

    private static string FirstString(JsonElement item, params string[] names)
    {
        if (item.ValueKind != JsonValueKind.Object)
            return string.Empty;

        foreach (var name in names)
        {
            if (!item.TryGetProperty(name, out var value))
                continue;

            if (value.ValueKind == JsonValueKind.String)
                return value.GetString() ?? string.Empty;

            if (value.ValueKind == JsonValueKind.Number || value.ValueKind == JsonValueKind.True || value.ValueKind == JsonValueKind.False)
                return value.ToString();
        }

        return string.Empty;
    }

    private (int created, int skipped) Fail(string error)
    {
        LastError = error;
        _logger.Error($"[TicketSystem] {error}");
        return (0, 0);
    }
}
