using Microsoft.Data.Sqlite;
using TaskStatus = TaskTool.Models.TaskStatus;
using TaskTool.Models;

namespace TaskTool.Services;

public class TaskService
{
    private readonly DatabaseService _db;
    private readonly LoggerService _logger;
    private readonly OutlookInteropService _outlook;

    public string LastError { get; private set; } = string.Empty;

    public TaskService(DatabaseService db, LoggerService logger, OutlookInteropService outlook, SettingsService settings)
    {
        _db = db;
        _logger = logger;
        _outlook = outlook;
    }

    public List<TaskItem> GetTasksForDay(DateTime day)
    {
        var list = new List<TaskItem>();
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT * FROM tasks WHERE date(start_local)=date($d) OR start_local IS NULL ORDER BY start_local";
        cmd.Parameters.AddWithValue("$d", day.ToString("yyyy-MM-dd"));
        using var reader = cmd.ExecuteReader();
        while (reader.Read()) list.Add(MapTask(reader));
        return list;
    }


    public List<TaskItem> GetAllTasks()
    {
        var list = new List<TaskItem>();
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT * FROM tasks ORDER BY CASE WHEN status='Running' THEN 0 WHEN status='Planned' THEN 1 WHEN status='Done' THEN 2 ELSE 3 END, COALESCE(start_local, created_utc) DESC";
        using var reader = cmd.ExecuteReader();
        while (reader.Read()) list.Add(MapTask(reader));
        return list;
    }
    public List<TaskItem> GetTasksForRange(DateTime from, DateTime to)
    {
        var list = new List<TaskItem>();
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT * FROM tasks WHERE start_local IS NOT NULL AND datetime(start_local)>=datetime($f) AND datetime(start_local)<datetime($t) ORDER BY start_local";
        cmd.Parameters.AddWithValue("$f", from.ToString("s"));
        cmd.Parameters.AddWithValue("$t", to.ToString("s"));
        using var reader = cmd.ExecuteReader();
        while (reader.Read()) list.Add(MapTask(reader));
        return list;
    }

    public TaskItem CreateTask(TaskItem task)
    {
        task.UpdatedUtc = DateTime.UtcNow;
        task.CreatedUtc = DateTime.UtcNow;
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"INSERT INTO tasks (id,title,description,ticket_url,start_local,end_local,status,priority,tags,outlook_entry_id,ticket_minutes_booked,created_utc,updated_utc)
VALUES ($id,$title,$desc,$url,$start,$end,$status,$priority,$tags,$entry,$ticket,$created,$updated)";
        BindTask(cmd, task);
        cmd.ExecuteNonQuery();
        return task;
    }

    public void UpdateTask(TaskItem task)
    {
        task.UpdatedUtc = DateTime.UtcNow;
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"UPDATE tasks SET title=$title,description=$desc,ticket_url=$url,start_local=$start,end_local=$end,status=$status,priority=$priority,tags=$tags,outlook_entry_id=$entry,ticket_minutes_booked=$ticket,updated_utc=$updated WHERE id=$id";
        BindTask(cmd, task);
        cmd.ExecuteNonQuery();
    }

    public void DeleteTask(TaskItem task)
    {
        var segments = GetSegments(task.Id);
        foreach (var segment in segments)
        {
            if (!DeleteSegmentOutlook(segment) && string.IsNullOrWhiteSpace(LastError) == false)
            {
                _logger.Warn($"Segment outlook delete failed for segment {segment.Id}: {LastError}");
            }
        }

        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();

        using (var segCmd = conn.CreateCommand())
        {
            segCmd.CommandText = "DELETE FROM task_segments WHERE task_id=$id";
            segCmd.Parameters.AddWithValue("$id", task.Id.ToString());
            segCmd.ExecuteNonQuery();
        }

        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM tasks WHERE id=$id";
        cmd.Parameters.AddWithValue("$id", task.Id.ToString());
        cmd.ExecuteNonQuery();
    }

    public void MarkDone(TaskItem task)
    {
        task.Status = TaskStatus.Done;
        UpdateTask(task);
    }

    public void MarkPlanned(TaskItem task)
    {
        task.Status = TaskStatus.Planned;
        UpdateTask(task);
    }

    public void StartTimer(TaskItem task)
    {
        task.Status = TaskStatus.Running;
        UpdateTask(task);
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "INSERT INTO time_logs (task_id,start_utc,note) VALUES ($id,$s,$n)";
        cmd.Parameters.AddWithValue("$id", task.Id.ToString());
        cmd.Parameters.AddWithValue("$s", DateTime.UtcNow.ToString("O"));
        cmd.Parameters.AddWithValue("$n", "running");
        cmd.ExecuteNonQuery();
    }

    public void PauseTimer(TaskItem task)
    {
        StopOpenLog(task.Id, "pause");
        task.Status = TaskStatus.Planned;
        UpdateTask(task);
    }

    public void StopTimer(TaskItem task)
    {
        StopOpenLog(task.Id, "stop");
        if (task.Status == TaskStatus.Running)
        {
            task.Status = TaskStatus.Planned;
            UpdateTask(task);
        }
    }

    public void AddTicketMinutes(TaskItem task, int minutes)
    {
        task.TicketMinutesBooked += minutes;
        UpdateTask(task);
    }

    public TimeSpan GetTrackedDuration(Guid taskId)
    {
        var total = TimeSpan.Zero;
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT start_utc,end_utc FROM time_logs WHERE task_id=$id";
        cmd.Parameters.AddWithValue("$id", taskId.ToString());
        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            var start = DateTime.Parse(reader["start_utc"].ToString()!).ToUniversalTime();
            var end = DateTime.TryParse(reader["end_utc"]?.ToString(), out var parsedEnd)
                ? parsedEnd.ToUniversalTime()
                : DateTime.UtcNow;
            if (end > start) total += end - start;
        }
        return total;
    }

    public int GetMonthTicketMinutes(DateTime month)
    {
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT COALESCE(SUM(ticket_minutes_booked),0) FROM tasks WHERE strftime('%Y-%m', COALESCE(start_local, created_utc)) = $m";
        cmd.Parameters.AddWithValue("$m", month.ToString("yyyy-MM"));
        return Convert.ToInt32(cmd.ExecuteScalar());
    }

    public List<(string Title, int Minutes)> GetTopTasksForMonth(DateTime month, int max = 5)
    {
        var result = new List<(string Title, int Minutes)>();
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"SELECT title, ticket_minutes_booked FROM tasks
WHERE strftime('%Y-%m', COALESCE(start_local, created_utc)) = $m
ORDER BY ticket_minutes_booked DESC, title ASC LIMIT $max";
        cmd.Parameters.AddWithValue("$m", month.ToString("yyyy-MM"));
        cmd.Parameters.AddWithValue("$max", max);

        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            result.Add((reader["title"]?.ToString() ?? "(ohne Titel)", Convert.ToInt32(reader["ticket_minutes_booked"])));
        }
        return result;
    }

    public TaskItem ParseQuickAdd(string input)
    {
        var parts = input.Split('|', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        var task = new TaskItem { Title = parts.ElementAtOrDefault(0) ?? "Neue Aufgabe" };
        if (DateTime.TryParse(parts.ElementAtOrDefault(1), out var start)) task.StartLocal = start;
        if (TryParseDuration(parts.ElementAtOrDefault(2), out var duration) && task.StartLocal.HasValue)
            task.EndLocal = task.StartLocal.Value.Add(duration);
        if (parts.Length > 3) task.TicketUrl = parts[3];
        return task;
    }

    public List<TaskItem> GetUpcomingTasks(DateTime fromInclusive, DateTime toExclusive)
    {
        return GetTasksForRange(fromInclusive, toExclusive)
            .Where(t => t.Status != TaskStatus.Done && t.Status != TaskStatus.Cancelled)
            .ToList();
    }

    public bool TestOutlookConnection()
    {
        LastError = string.Empty;
        var result = _outlook.TestConnection();
        if (!result.ok)
        {
            LastError = $"Outlook Verbindungstest fehlgeschlagen: {result.error}";
            return false;
        }

        return true;
    }

    public List<TaskSegment> GetSegments(Guid taskId)
    {
        var list = new List<TaskSegment>();
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT id,task_id,start_local,end_local,planned_minutes,note,outlook_entry_id FROM task_segments WHERE task_id=$id ORDER BY start_local";
        cmd.Parameters.AddWithValue("$id", taskId.ToString());
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            list.Add(new TaskSegment
            {
                Id = Convert.ToInt64(r["id"]),
                TaskId = Guid.Parse(r["task_id"].ToString()!),
                StartLocal = DateTime.Parse(r["start_local"].ToString()!),
                EndLocal = DateTime.Parse(r["end_local"].ToString()!),
                PlannedMinutes = Convert.ToInt32(r["planned_minutes"]),
                Note = r["note"]?.ToString() ?? string.Empty,
                OutlookEntryId = r["outlook_entry_id"]?.ToString() ?? string.Empty
            });
        }
        return list;
    }

    public void AddSegment(TaskSegment segment)
    {
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "INSERT INTO task_segments(task_id,start_local,end_local,planned_minutes,note,outlook_entry_id) VALUES ($t,$s,$e,$p,$n,$o)";
        cmd.Parameters.AddWithValue("$t", segment.TaskId.ToString());
        cmd.Parameters.AddWithValue("$s", segment.StartLocal.ToString("s"));
        cmd.Parameters.AddWithValue("$e", segment.EndLocal.ToString("s"));
        cmd.Parameters.AddWithValue("$p", (int)(segment.EndLocal - segment.StartLocal).TotalMinutes);
        cmd.Parameters.AddWithValue("$n", segment.Note);
        cmd.Parameters.AddWithValue("$o", segment.OutlookEntryId);
        cmd.ExecuteNonQuery();

        using var idCmd = conn.CreateCommand();
        idCmd.CommandText = "SELECT last_insert_rowid()";
        segment.Id = Convert.ToInt64(idCmd.ExecuteScalar());
    }

    public void UpdateSegment(TaskSegment segment)
    {
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE task_segments SET start_local=$s,end_local=$e,planned_minutes=$p,note=$n WHERE id=$id";
        cmd.Parameters.AddWithValue("$s", segment.StartLocal.ToString("s"));
        cmd.Parameters.AddWithValue("$e", segment.EndLocal.ToString("s"));
        cmd.Parameters.AddWithValue("$p", (int)(segment.EndLocal - segment.StartLocal).TotalMinutes);
        cmd.Parameters.AddWithValue("$n", segment.Note);
        cmd.Parameters.AddWithValue("$id", segment.Id);
        cmd.ExecuteNonQuery();
    }

    public void DeleteSegment(long segmentId)
    {
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM task_segments WHERE id=$id";
        cmd.Parameters.AddWithValue("$id", segmentId);
        cmd.ExecuteNonQuery();
    }

    public bool SyncSegmentOutlook(TaskSegment segment, string title, string description, string ticketUrl)
    {
        LastError = string.Empty;
        var body = $"{description}\n{ticketUrl}\nTaskID: {segment.TaskId}\nSegmentID: {segment.Id}\nNotiz: {segment.Note}";
        var result = _outlook.UpsertBlock(segment.OutlookEntryId, title, body, segment.StartLocal, segment.EndLocal);
        if (!result.ok)
        {
            LastError = $"Outlook Sync Fehler: {result.error}";
            return false;
        }

        segment.OutlookEntryId = result.entryId;
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE task_segments SET outlook_entry_id=$o WHERE id=$id";
        cmd.Parameters.AddWithValue("$o", segment.OutlookEntryId);
        cmd.Parameters.AddWithValue("$id", segment.Id);
        cmd.ExecuteNonQuery();
        return true;
    }

    public bool DeleteSegmentOutlook(TaskSegment segment)
    {
        if (string.IsNullOrWhiteSpace(segment.OutlookEntryId)) return true;
        var result = _outlook.DeleteBlock(segment.OutlookEntryId);
        if (!result.ok)
        {
            LastError = result.error;
            return false;
        }

        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE task_segments SET outlook_entry_id='' WHERE id=$id";
        cmd.Parameters.AddWithValue("$id", segment.Id);
        cmd.ExecuteNonQuery();
        segment.OutlookEntryId = string.Empty;
        return true;
    }

    private bool TryParseDuration(string? text, out TimeSpan duration)
    {
        duration = TimeSpan.Zero;
        if (string.IsNullOrWhiteSpace(text)) return false;
        text = text.Trim().ToLowerInvariant();
        if (text.EndsWith("m") && int.TryParse(text[..^1], out var mins))
        {
            duration = TimeSpan.FromMinutes(mins);
            return true;
        }
        if (text.EndsWith("h") && int.TryParse(text[..^1], out var h))
        {
            duration = TimeSpan.FromHours(h);
            return true;
        }
        return TimeSpan.TryParse(text, out duration);
    }

    private void StopOpenLog(Guid taskId, string note)
    {
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE time_logs SET end_utc=$e,note=$n WHERE id=(SELECT id FROM time_logs WHERE task_id=$id AND end_utc IS NULL ORDER BY id DESC LIMIT 1)";
        cmd.Parameters.AddWithValue("$e", DateTime.UtcNow.ToString("O"));
        cmd.Parameters.AddWithValue("$n", note);
        cmd.Parameters.AddWithValue("$id", taskId.ToString());
        cmd.ExecuteNonQuery();
    }

    private static TaskItem MapTask(SqliteDataReader reader)
    {
        return new TaskItem
        {
            Id = Guid.Parse(reader.GetString(reader.GetOrdinal("id"))),
            Title = reader.GetString(reader.GetOrdinal("title")),
            Description = reader["description"]?.ToString() ?? string.Empty,
            TicketUrl = reader["ticket_url"]?.ToString() ?? string.Empty,
            StartLocal = DateTime.TryParse(reader["start_local"]?.ToString(), out var s) ? s : null,
            EndLocal = DateTime.TryParse(reader["end_local"]?.ToString(), out var e) ? e : null,
            Status = Enum.TryParse<TaskStatus>(reader["status"]?.ToString(), out var st) ? st : TaskStatus.Planned,
            Priority = reader["priority"] == DBNull.Value ? null : Convert.ToInt32(reader["priority"]),
            Tags = reader["tags"]?.ToString() ?? string.Empty,
            OutlookEntryId = reader["outlook_entry_id"]?.ToString() ?? string.Empty,
            TicketMinutesBooked = Convert.ToInt32(reader["ticket_minutes_booked"]),
            CreatedUtc = DateTime.Parse(reader["created_utc"].ToString()!),
            UpdatedUtc = DateTime.Parse(reader["updated_utc"].ToString()!)
        };
    }

    private static void BindTask(SqliteCommand cmd, TaskItem task)
    {
        cmd.Parameters.AddWithValue("$id", task.Id.ToString());
        cmd.Parameters.AddWithValue("$title", task.Title);
        cmd.Parameters.AddWithValue("$desc", task.Description);
        cmd.Parameters.AddWithValue("$url", task.TicketUrl);
        cmd.Parameters.AddWithValue("$start", task.StartLocal?.ToString("s") ?? (object)DBNull.Value);
        cmd.Parameters.AddWithValue("$end", task.EndLocal?.ToString("s") ?? (object)DBNull.Value);
        cmd.Parameters.AddWithValue("$status", task.Status.ToString());
        cmd.Parameters.AddWithValue("$priority", task.Priority ?? (object)DBNull.Value);
        cmd.Parameters.AddWithValue("$tags", task.Tags);
        cmd.Parameters.AddWithValue("$entry", task.OutlookEntryId);
        cmd.Parameters.AddWithValue("$ticket", task.TicketMinutesBooked);
        cmd.Parameters.AddWithValue("$created", task.CreatedUtc.ToString("O"));
        cmd.Parameters.AddWithValue("$updated", task.UpdatedUtc.ToString("O"));
    }
}
