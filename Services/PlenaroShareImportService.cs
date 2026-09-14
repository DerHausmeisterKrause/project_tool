using Microsoft.Data.Sqlite;
using TaskTool.Models;

namespace TaskTool.Services;

/// Imports Outlook snapshots only; it deliberately never calls Znuny or writes Outlook.
public sealed class PlenaroShareImportService
{
    private readonly DatabaseService _db;
    private readonly TaskService _tasks;
    private readonly SettingsService _settings;
    private readonly LoggerService _logger;

    public PlenaroShareImportService(DatabaseService db, TaskService tasks, SettingsService settings, LoggerService logger)
    {
        _db = db;
        _tasks = tasks;
        _settings = settings;
        _logger = logger;
    }

    public void Import(IEnumerable<OutlookCalendarEvent> events)
    {
        foreach (var calendarEvent in events)
        {
            try
            {
                if (!PlenaroShareCodec.TryParse(calendarEvent.FullBody, out var payload, out var hash, out var reason))
                {
                    if (calendarEvent.FullBody.Contains(PlenaroShareCodec.BeginMarker, StringComparison.Ordinal))
                        _logger.Warning($"[PlenaroShareImport] action=skipped reason={reason}");
                    continue;
                }
                if (string.Equals(payload!.OriginClientInstanceId, _settings.Current.ClientInstanceId, StringComparison.OrdinalIgnoreCase))
                {
                    _logger.Info("[PlenaroShareImport] action=skipped reason=self-origin");
                    continue;
                }
                ImportOne(payload, hash);
            }
            catch (Exception ex)
            {
                _logger.Warning($"[PlenaroShareImport] action=skipped reason=event-error type={ex.GetType().Name}");
            }
        }
    }

    private void ImportOne(PlenaroSharePayloadV1 payload, string hash)
    {
        using var connection = new SqliteConnection(_db.ConnectionString);
        connection.Open();

        Guid mappedTaskId = Guid.Empty;
        DateTime? lastGeneratedUtc = null;
        using (var command = connection.CreateCommand())
        {
            command.CommandText = "SELECT local_task_id,last_generated_utc FROM plenaro_shared_task_imports WHERE task_share_id=$id";
            command.Parameters.AddWithValue("$id", payload.TaskShareId);
            using var reader = command.ExecuteReader();
            if (reader.Read())
            {
                Guid.TryParse(reader.GetString(0), out mappedTaskId);
                if (!reader.IsDBNull(1) && DateTime.TryParse(reader.GetString(1), null,
                        System.Globalization.DateTimeStyles.RoundtripKind, out var parsed))
                    lastGeneratedUtc = parsed.ToUniversalTime();
            }
        }

        long mappedSegmentId = 0;
        string segmentHash = string.Empty;
        using (var command = connection.CreateCommand())
        {
            command.CommandText = "SELECT local_segment_id,last_payload_hash FROM plenaro_shared_segment_imports WHERE segment_share_id=$id";
            command.Parameters.AddWithValue("$id", payload.SegmentShareId);
            using var reader = command.ExecuteReader();
            if (reader.Read())
            {
                mappedSegmentId = reader.GetInt64(0);
                segmentHash = reader.GetString(1);
            }
        }

        var allTasks = _tasks.GetAllTasks();
        var task = mappedTaskId == Guid.Empty ? null : allTasks.SingleOrDefault(item => item.Id == mappedTaskId);
        if (task == null)
        {
            var sharedMatches = allTasks.Where(item => string.Equals(item.TaskShareId, payload.TaskShareId, StringComparison.OrdinalIgnoreCase)).ToList();
            if (sharedMatches.Count == 1)
                task = sharedMatches[0];
        }
        if (task == null && !string.IsNullOrWhiteSpace(payload.TicketId))
        {
            var ticketMatches = allTasks.Where(item => HasTicketId(item.Tags, payload.TicketId)).ToList();
            if (ticketMatches.Count == 1)
                task = ticketMatches[0];
        }

        var existingSegment = task == null || mappedSegmentId == 0
            ? null
            : _tasks.GetSegments(task.Id).SingleOrDefault(segment => segment.Id == mappedSegmentId);
        if (mappedTaskId != Guid.Empty && existingSegment != null && string.Equals(segmentHash, hash, StringComparison.Ordinal))
            return;

        var created = task == null;
        task ??= new TaskItem { Status = TaskTool.Models.TaskStatus.Planned, IsZnunyAssigned = false };
        var acceptsSnapshot = !lastGeneratedUtc.HasValue || payload.GeneratedUtc.ToUniversalTime() >= lastGeneratedUtc.Value;
        if (acceptsSnapshot)
        {
            task.Title = payload.TaskTitle;
            task.Description = payload.TaskDescription;
            task.TicketUrl = payload.TicketUrl;
            task.TicketState = payload.TicketState;
            if (!string.IsNullOrWhiteSpace(payload.TicketId))
                task.Tags = MergeTicketTags(task.Tags, payload.TicketId, payload.TicketNumber);
        }
        task.IsPlenaroShared = true;
        task.TaskShareId = payload.TaskShareId;
        task.ShareOriginClientInstanceId = payload.OriginClientInstanceId;
        if (created)
            _tasks.CreateTask(task, true);
        else if (acceptsSnapshot || mappedTaskId == Guid.Empty)
            _tasks.UpdateTask(task, false);

        using (var command = connection.CreateCommand())
        {
            command.CommandText = @"INSERT INTO plenaro_shared_task_imports(task_share_id,local_task_id,origin_client_instance_id,last_payload_hash,last_generated_utc)
VALUES($share,$local,$origin,$hash,$generated)
ON CONFLICT(task_share_id) DO UPDATE SET local_task_id=$local,origin_client_instance_id=$origin,
last_payload_hash=$hash,last_generated_utc=CASE WHEN $accept=1 THEN $generated ELSE last_generated_utc END";
            command.Parameters.AddWithValue("$share", payload.TaskShareId);
            command.Parameters.AddWithValue("$local", task.Id.ToString());
            command.Parameters.AddWithValue("$origin", payload.OriginClientInstanceId);
            command.Parameters.AddWithValue("$hash", hash); // retained for V29 compatibility; never used for deduplication
            command.Parameters.AddWithValue("$generated", payload.GeneratedUtc.ToUniversalTime().ToString("O"));
            command.Parameters.AddWithValue("$accept", acceptsSnapshot ? 1 : 0);
            command.ExecuteNonQuery();
        }

        var segment = existingSegment ?? new TaskSegment { TaskId = task.Id, SegmentShareId = payload.SegmentShareId, IsSharedImport = true };
        segment.StartLocal = payload.SegmentStartLocal;
        segment.EndLocal = payload.SegmentEndLocal;
        segment.Note = payload.SegmentNote;
        segment.PlannedMinutes = (int)(segment.EndLocal - segment.StartLocal).TotalMinutes;
        if (segment.Id == 0)
            _tasks.AddSegment(segment);
        else
            _tasks.UpdateSegment(segment);

        // This is the success marker. It deliberately comes after the segment write so a
        // partial task import is retried on the next calendar scan.
        using (var command = connection.CreateCommand())
        {
            command.CommandText = @"INSERT INTO plenaro_shared_segment_imports(segment_share_id,local_segment_id,task_share_id,last_payload_hash)
VALUES($share,$local,$task,$hash)
ON CONFLICT(segment_share_id) DO UPDATE SET local_segment_id=$local,task_share_id=$task,last_payload_hash=$hash";
            command.Parameters.AddWithValue("$share", payload.SegmentShareId);
            command.Parameters.AddWithValue("$local", segment.Id);
            command.Parameters.AddWithValue("$task", payload.TaskShareId);
            command.Parameters.AddWithValue("$hash", hash);
            command.ExecuteNonQuery();
        }
        _logger.Info($"[PlenaroShareImport] action={(created ? "import-created" : "import-updated")} localTaskId={task.Id} hasTicket={!string.IsNullOrWhiteSpace(payload.TicketId)}");
    }

    private static bool HasTicketId(string tags, string id) => tags.Split(';', StringSplitOptions.TrimEntries)
        .Any(value => string.Equals(value, $"ZnunyTicketID:{id}", StringComparison.OrdinalIgnoreCase));

    private static string MergeTicketTags(string tags, string id, string number)
    {
        var values = tags.Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(value => !value.StartsWith("ZnunyTicketID:", StringComparison.OrdinalIgnoreCase)
                            && !value.StartsWith("ZnunyTicketNumber:", StringComparison.OrdinalIgnoreCase)).ToList();
        values.Add($"ZnunyTicketID:{id}");
        if (!string.IsNullOrWhiteSpace(number)) values.Add($"ZnunyTicketNumber:{number}");
        return string.Join(';', values);
    }
}
