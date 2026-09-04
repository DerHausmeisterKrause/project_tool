using Microsoft.Data.Sqlite;
using System.Globalization;
using TaskStatus = TaskTool.Models.TaskStatus;
using TaskTool.Models;

namespace TaskTool.Services;

public class TaskService
{
    private readonly DatabaseService _db;
    private readonly LoggerService _logger;
    private readonly OutlookInteropService _outlook;

    public string LastError { get; private set; } = string.Empty;
    public event Action? SegmentsChanged;

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

    public TaskItem CreateTask(TaskItem task, bool isBackgroundImport = false)
    {
        task.UpdatedUtc = DateTime.UtcNow;
        task.CreatedUtc = DateTime.UtcNow;
        task.LocalActivityUtc = isBackgroundImport
            ? task.TicketCreatedUtc ?? task.CreatedUtc
            : task.CreatedUtc;
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"INSERT INTO tasks (id,title,description,ticket_url,start_local,end_local,status,priority,tags,outlook_entry_id,ticket_minutes_booked,ticket_seconds_booked,is_pinned,is_znuny_assigned,created_utc,updated_utc,ticket_created_utc,ticket_changed_utc,local_activity_utc,ticket_state,ticket_state_type)
VALUES ($id,$title,$desc,$url,$start,$end,$status,$priority,$tags,$entry,$ticket,$ticketSeconds,$pinned,$znunyAssigned,$created,$updated,$ticketCreated,$ticketChanged,$localActivity,$ticketState,$ticketStateType)";
        BindTask(cmd, task);
        cmd.ExecuteNonQuery();
        return task;
    }

    public void UpdateTask(TaskItem task, bool touchLocalActivity = true)
    {
        task.UpdatedUtc = DateTime.UtcNow;
        if (touchLocalActivity)
            task.LocalActivityUtc = task.UpdatedUtc;
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"UPDATE tasks SET title=$title,description=$desc,ticket_url=$url,start_local=$start,end_local=$end,status=$status,priority=$priority,tags=$tags,outlook_entry_id=$entry,ticket_minutes_booked=$ticket,ticket_seconds_booked=$ticketSeconds,is_pinned=$pinned,is_znuny_assigned=$znunyAssigned,updated_utc=$updated,ticket_created_utc=$ticketCreated,ticket_changed_utc=$ticketChanged,local_activity_utc=$localActivity,ticket_state=$ticketState,ticket_state_type=$ticketStateType WHERE id=$id";
        BindTask(cmd, task);
        cmd.ExecuteNonQuery();
    }

    public void TouchTaskActivity(Guid taskId)
    {
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE tasks SET local_activity_utc=$activity WHERE id=$id";
        cmd.Parameters.AddWithValue("$activity", DateTime.UtcNow.ToString("O"));
        cmd.Parameters.AddWithValue("$id", taskId.ToString());
        cmd.ExecuteNonQuery();
    }

    public void DeleteTask(TaskItem task)
    {
        var segments = GetSegments(task.Id);
        foreach (var segment in segments)
        {
            if (!DeleteSegmentOutlook(segment) && string.IsNullOrWhiteSpace(LastError) == false)
            {
                _logger.Error($"Segment outlook delete failed for segment {segment.Id}: {LastError}");
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

    public void SetPinned(TaskItem task, bool isPinned)
    {
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE tasks SET is_pinned=$pinned WHERE id=$id";
        cmd.Parameters.AddWithValue("$pinned", isPinned ? 1 : 0);
        cmd.Parameters.AddWithValue("$id", task.Id.ToString());
        cmd.ExecuteNonQuery();
        task.IsPinned = isPinned;
    }

    public void SetZnunyAssigned(TaskItem task, bool assigned)
    {
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE tasks SET is_znuny_assigned=$assigned WHERE id=$id";
        cmd.Parameters.AddWithValue("$assigned", assigned ? 1 : 0);
        cmd.Parameters.AddWithValue("$id", task.Id.ToString());
        cmd.ExecuteNonQuery();
        task.IsZnunyAssigned = assigned;
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
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();

        using (var closeOthers = conn.CreateCommand())
        {
            closeOthers.CommandText = "UPDATE time_logs SET end_utc=$e,note='auto-stop' WHERE end_utc IS NULL AND task_id <> $id";
            closeOthers.Parameters.AddWithValue("$e", DateTime.UtcNow.ToString("O"));
            closeOthers.Parameters.AddWithValue("$id", task.Id.ToString());
            closeOthers.ExecuteNonQuery();
        }

        using (var resetStatuses = conn.CreateCommand())
        {
            resetStatuses.CommandText = "UPDATE tasks SET status=$planned WHERE status=$running AND id<>$id";
            resetStatuses.Parameters.AddWithValue("$planned", TaskStatus.Planned.ToString());
            resetStatuses.Parameters.AddWithValue("$running", TaskStatus.Running.ToString());
            resetStatuses.Parameters.AddWithValue("$id", task.Id.ToString());
            resetStatuses.ExecuteNonQuery();
        }

        task.Status = TaskStatus.Running;
        UpdateTask(task);

        using var cmd = conn.CreateCommand();
        cmd.CommandText = "INSERT INTO time_logs (task_id,start_utc,note) VALUES ($id,$s,$n)";
        cmd.Parameters.AddWithValue("$id", task.Id.ToString());
        cmd.Parameters.AddWithValue("$s", DateTime.UtcNow.ToString("O"));
        cmd.Parameters.AddWithValue("$n", "running");
        cmd.ExecuteNonQuery();
    }

    public void PauseTimer(TaskItem task)
    {
        var elapsedSeconds = CloseOpenLogAndGetElapsedSeconds(task.Id, "pause");
        if (elapsedSeconds > 0)
            task.TicketSecondsBooked = Math.Max(0, task.TicketSecondsBooked + elapsedSeconds);

        task.TicketMinutesBooked = (int)(task.TicketSecondsBooked / 60);

        task.Status = TaskStatus.Planned;
        UpdateTask(task);
    }

    public void StopTimer(TaskItem task)
    {
        var elapsedSeconds = CloseOpenLogAndGetElapsedSeconds(task.Id, "stop");
        if (elapsedSeconds > 0)
            task.TicketSecondsBooked = Math.Max(0, task.TicketSecondsBooked + elapsedSeconds);

        task.TicketMinutesBooked = (int)(task.TicketSecondsBooked / 60);

        if (task.Status == TaskStatus.Running)
            task.Status = TaskStatus.Planned;

        UpdateTask(task);
    }

    public void AddTicketMinutes(TaskItem task, int minutes)
    {
        task.TicketSecondsBooked = Math.Max(0, task.TicketSecondsBooked + (minutes * 60L));
        task.TicketMinutesBooked = (int)(task.TicketSecondsBooked / 60);
        UpdateTask(task);
    }

    public void SetUnbookedTicketSeconds(TaskItem task, long desiredUnbookedSeconds,
        long bookingBaselineSeconds, long successfullyTransferredSeconds)
    {
        ArgumentOutOfRangeException.ThrowIfNegative(desiredUnbookedSeconds);
        var wasRunning = task.Status == TaskStatus.Running;
        if (wasRunning) StopTimer(task);
        task.TicketSecondsBooked = checked(Math.Max(0, bookingBaselineSeconds)
            + Math.Max(0, successfullyTransferredSeconds) + desiredUnbookedSeconds);
        task.TicketMinutesBooked = (int)(task.TicketSecondsBooked / 60);
        UpdateTask(task);
        if (wasRunning) StartTimer(task);
    }

    public void CreateTicketTimeBooking(TicketTimeBooking booking)
    {
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"INSERT INTO ticket_time_bookings
(id,task_id,ticket_id,ticket_number,booking_id,article_id,minutes,booked_minutes,source_seconds,booked_at_utc,short_description,note,cost_center,order_value,status)
VALUES ($id,$task,$ticket,$number,$booking,$article,$minutes,$bookedMinutes,$seconds,$booked,$description,$note,$cost,$order,$status)";
        BindTicketTimeBooking(cmd, booking);
        cmd.ExecuteNonQuery();
    }

    public void CompleteTicketTimeBooking(TicketTimeBooking booking, string articleId)
    {
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"UPDATE ticket_time_bookings
SET article_id=$article,
    status='Succeeded',
    booked_at_utc=$booked,
    minutes=booked_minutes
WHERE id=$id AND status<>'Succeeded'";
        cmd.Parameters.AddWithValue("$article", articleId ?? string.Empty);
        cmd.Parameters.AddWithValue("$booked", DateTime.UtcNow.ToString("O"));
        cmd.Parameters.AddWithValue("$id", booking.Id.ToString());
        if (cmd.ExecuteNonQuery() > 0)
        {
            booking.Minutes = booking.BookedMinutes;
            TouchTaskActivity(booking.TaskId);
        }
    }

    public void FailTicketTimeBooking(TicketTimeBooking booking)
    {
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE ticket_time_bookings SET status='Failed' WHERE id=$id AND status='Pending'";
        cmd.Parameters.AddWithValue("$id", booking.Id.ToString());
        cmd.ExecuteNonQuery();
    }

    public void ResetTicketTimeBookingForRetry(TicketTimeBooking booking)
    {
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE ticket_time_bookings SET status='Pending' WHERE id=$id AND status='Failed'";
        cmd.Parameters.AddWithValue("$id", booking.Id.ToString());
        cmd.ExecuteNonQuery();
    }

    public TicketTimeBooking? GetPendingTicketTimeBooking(Guid taskId)
        => GetTicketTimeBookings(taskId, "Pending").FirstOrDefault();

    public List<TicketTimeBooking> GetSuccessfulTicketTimeBookings(Guid taskId)
        => GetTicketTimeBookings(taskId, "Succeeded");

    public List<TicketTimeBooking> GetAllTicketTimeBookings(Guid taskId)
        => GetTicketTimeBookings(taskId, null);

    public long GetSuccessfullyTransferredSeconds(Guid taskId)
        => GetSuccessfulTicketTimeBookings(taskId).Sum(booking => booking.SourceSeconds);

    public long GetTicketTimeBookingBaselineSeconds(Guid taskId)
    {
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT baseline_seconds FROM ticket_time_booking_baselines WHERE task_id=$task";
        cmd.Parameters.AddWithValue("$task", taskId.ToString());
        return Convert.ToInt64(cmd.ExecuteScalar() ?? 0L);
    }

    private List<TicketTimeBooking> GetTicketTimeBookings(Guid taskId, string? status)
    {
        var result = new List<TicketTimeBooking>();
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT * FROM ticket_time_bookings WHERE task_id=$task AND ($status IS NULL OR status=$status) ORDER BY booked_at_utc DESC";
        cmd.Parameters.AddWithValue("$task", taskId.ToString());
        cmd.Parameters.AddWithValue("$status", (object?)status ?? DBNull.Value);
        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new TicketTimeBooking
            {
                Id = Guid.Parse(reader["id"].ToString()!),
                TaskId = Guid.Parse(reader["task_id"].ToString()!),
                TicketId = reader["ticket_id"].ToString() ?? string.Empty,
                TicketNumber = reader["ticket_number"].ToString() ?? string.Empty,
                BookingId = reader["booking_id"].ToString() ?? string.Empty,
                ArticleId = reader["article_id"].ToString() ?? string.Empty,
                Minutes = Convert.ToDecimal(reader["minutes"]),
                BookedMinutes = Convert.ToDecimal(reader["booked_minutes"]),
                SourceSeconds = Convert.ToInt64(reader["source_seconds"]),
                BookedAtUtc = DateTime.Parse(reader["booked_at_utc"].ToString()!, null, System.Globalization.DateTimeStyles.RoundtripKind),
                ShortDescription = reader["short_description"].ToString() ?? string.Empty,
                Note = reader["note"].ToString() ?? string.Empty,
                CostCenter = reader["cost_center"].ToString() ?? string.Empty,
                Order = reader["order_value"].ToString() ?? string.Empty,
                Status = reader["status"].ToString() ?? string.Empty
            });
        }
        return result;
    }

    private static void BindTicketTimeBooking(SqliteCommand cmd, TicketTimeBooking booking)
    {
        cmd.Parameters.AddWithValue("$id", booking.Id.ToString());
        cmd.Parameters.AddWithValue("$task", booking.TaskId.ToString());
        cmd.Parameters.AddWithValue("$ticket", booking.TicketId);
        cmd.Parameters.AddWithValue("$number", booking.TicketNumber);
        cmd.Parameters.AddWithValue("$booking", booking.BookingId);
        cmd.Parameters.AddWithValue("$article", booking.ArticleId);
        cmd.Parameters.AddWithValue("$minutes", Convert.ToDouble(booking.Minutes));
        cmd.Parameters.AddWithValue("$bookedMinutes", Convert.ToDouble(booking.BookedMinutes));
        cmd.Parameters.AddWithValue("$seconds", booking.SourceSeconds);
        cmd.Parameters.AddWithValue("$booked", booking.BookedAtUtc.ToString("O"));
        cmd.Parameters.AddWithValue("$description", booking.ShortDescription);
        cmd.Parameters.AddWithValue("$note", booking.Note);
        cmd.Parameters.AddWithValue("$cost", booking.CostCenter);
        cmd.Parameters.AddWithValue("$order", booking.Order);
        cmd.Parameters.AddWithValue("$status", booking.Status);
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
            var start = ParseRequiredDateTime(reader["start_utc"].ToString()).ToUniversalTime();
            var end = DateTime.TryParse(reader["end_utc"]?.ToString(), out var parsedEnd)
                ? parsedEnd.ToUniversalTime()
                : DateTime.UtcNow;
            if (end > start) total += end - start;
        }
        return total;
    }


    public TimeSpan GetOpenSessionDuration(Guid taskId)
    {
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT start_utc FROM time_logs WHERE task_id=$id AND end_utc IS NULL ORDER BY id DESC LIMIT 1";
        cmd.Parameters.AddWithValue("$id", taskId.ToString());
        var value = cmd.ExecuteScalar()?.ToString();
        if (string.IsNullOrWhiteSpace(value) || !DateTime.TryParse(value, out var startUtc))
            return TimeSpan.Zero;

        var elapsed = DateTime.UtcNow - startUtc.ToUniversalTime();
        return elapsed > TimeSpan.Zero ? elapsed : TimeSpan.Zero;
    }

    public SuccessfulBookingStatistics GetSuccessfulBookingStatistics(
        DateTime localDay,
        TimeZoneInfo calendarTimeZone)
    {
        var secondsByMonth = new Dictionary<DateTime, long>();
        long todaySeconds = 0;
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"SELECT booked_at_utc, booked_minutes
FROM ticket_time_bookings
WHERE status = 'Succeeded'";
        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            if (!DateTimeOffset.TryParse(
                    reader["booked_at_utc"]?.ToString(),
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                    out var bookedAtUtc))
                continue;

            var bookedLocal = TimeZoneInfo.ConvertTime(bookedAtUtc, calendarTimeZone);
            var seconds = checked(Math.Max(0, Convert.ToInt64(Convert.ToDecimal(reader["booked_minutes"]) * 60m)));
            var month = new DateTime(bookedLocal.Year, bookedLocal.Month, 1);
            secondsByMonth[month] = secondsByMonth.GetValueOrDefault(month) + seconds;
            if (bookedLocal.Date == localDay.Date)
                todaySeconds += seconds;
        }

        return new SuccessfulBookingStatistics(todaySeconds, secondsByMonth);
    }

    // These task-based counters remain for the existing Today view. Reports use
    // GetSuccessfulBookingStatistics and never infer booking dates from tasks.
    public int GetTicketMinutesForDay(DateTime day)
        => GetTaskTicketMinutesForRange(day.Date, day.Date.AddDays(1));

    public int GetMonthTicketMinutes(DateTime month)
    {
        var monthStart = new DateTime(month.Year, month.Month, 1);
        return GetTaskTicketMinutesForRange(monthStart, monthStart.AddMonths(1));
    }

    private int GetTaskTicketMinutesForRange(DateTime fromInclusive, DateTime toExclusive)
    {
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"SELECT COALESCE(SUM(ticket_minutes_booked), 0)
FROM tasks
WHERE datetime(COALESCE(start_local, created_utc)) >= datetime($from)
  AND datetime(COALESCE(start_local, created_utc)) < datetime($to)";
        cmd.Parameters.AddWithValue("$from", fromInclusive.ToString("s"));
        cmd.Parameters.AddWithValue("$to", toExclusive.ToString("s"));
        return Convert.ToInt32(cmd.ExecuteScalar());
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


    public List<(TaskItem Task, TaskSegment Segment)> GetSegmentsForRange(DateTime fromInclusive, DateTime toExclusive)
    {
        var result = new List<(TaskItem Task, TaskSegment Segment)>();
        var tasks = GetAllTasks().ToDictionary(t => t.Id, t => t);

        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT id,task_id,start_local,end_local,planned_minutes,note,outlook_entry_id FROM task_segments WHERE datetime(start_local)>=datetime($from) AND datetime(start_local)<datetime($to) ORDER BY start_local";
        cmd.Parameters.AddWithValue("$from", fromInclusive.ToString("s"));
        cmd.Parameters.AddWithValue("$to", toExclusive.ToString("s"));
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var taskId = Guid.Parse(r["task_id"].ToString()!);
            if (!tasks.TryGetValue(taskId, out var task))
                continue;

            var segment = new TaskSegment
            {
                Id = Convert.ToInt64(r["id"]),
                TaskId = taskId,
                StartLocal = ParseRequiredDateTime(r["start_local"].ToString()),
                EndLocal = ParseRequiredDateTime(r["end_local"].ToString()),
                PlannedMinutes = Convert.ToInt32(r["planned_minutes"]),
                Note = r["note"]?.ToString() ?? string.Empty,
                OutlookEntryId = r["outlook_entry_id"]?.ToString() ?? string.Empty
            };

            result.Add((task, segment));
        }

        return result;
    }

    public HashSet<Guid> GetTaskIdsWithSegmentsForRange(DateTime fromInclusive, DateTime toExclusive)
    {
        var taskIds = new HashSet<Guid>();
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"SELECT DISTINCT task_id
FROM task_segments
WHERE datetime(start_local) >= datetime($from)
  AND datetime(start_local) < datetime($to)";
        cmd.Parameters.AddWithValue("$from", fromInclusive.ToString("s"));
        cmd.Parameters.AddWithValue("$to", toExclusive.ToString("s"));

        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            if (Guid.TryParse(reader["task_id"]?.ToString(), out var taskId))
                taskIds.Add(taskId);
        }

        return taskIds;
    }

    public HashSet<Guid> GetTaskIdsWithActiveOrFutureSegments(DateTime now)
    {
        var taskIds = new HashSet<Guid>();
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"SELECT DISTINCT task_id
FROM task_segments
WHERE datetime(end_local) > datetime($now)";
        cmd.Parameters.AddWithValue("$now", now.ToString("s"));
        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            if (Guid.TryParse(reader["task_id"]?.ToString(), out var taskId))
                taskIds.Add(taskId);
        }
        return taskIds;
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
                StartLocal = ParseRequiredDateTime(r["start_local"].ToString()),
                EndLocal = ParseRequiredDateTime(r["end_local"].ToString()),
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
        TouchTaskActivity(segment.TaskId);
        SegmentsChanged?.Invoke();
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
        if (cmd.ExecuteNonQuery() > 0)
        {
            TouchTaskActivity(segment.TaskId);
            SegmentsChanged?.Invoke();
        }
    }

    public void DeleteSegment(long segmentId)
    {
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        Guid? taskId = null;
        using (var lookup = conn.CreateCommand())
        {
            lookup.CommandText = "SELECT task_id FROM task_segments WHERE id=$id";
            lookup.Parameters.AddWithValue("$id", segmentId);
            if (Guid.TryParse(lookup.ExecuteScalar()?.ToString(), out var parsedTaskId))
                taskId = parsedTaskId;
        }
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM task_segments WHERE id=$id";
        cmd.Parameters.AddWithValue("$id", segmentId);
        if (cmd.ExecuteNonQuery() > 0)
        {
            if (taskId.HasValue)
                TouchTaskActivity(taskId.Value);
            SegmentsChanged?.Invoke();
        }
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

    private long CloseOpenLogAndGetElapsedSeconds(Guid taskId, string note)
    {
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();

        DateTime? startUtc = null;
        using (var read = conn.CreateCommand())
        {
            read.CommandText = "SELECT start_utc FROM time_logs WHERE task_id=$id AND end_utc IS NULL ORDER BY id DESC LIMIT 1";
            read.Parameters.AddWithValue("$id", taskId.ToString());
            var value = read.ExecuteScalar()?.ToString();
            if (DateTime.TryParse(value, out var parsed))
                startUtc = parsed.ToUniversalTime();
        }

        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = "UPDATE time_logs SET end_utc=$e,note=$n WHERE id=(SELECT id FROM time_logs WHERE task_id=$id AND end_utc IS NULL ORDER BY id DESC LIMIT 1)";
            cmd.Parameters.AddWithValue("$e", DateTime.UtcNow.ToString("O"));
            cmd.Parameters.AddWithValue("$n", note);
            cmd.Parameters.AddWithValue("$id", taskId.ToString());
            cmd.ExecuteNonQuery();
        }

        if (!startUtc.HasValue)
            return 0;

        var elapsed = DateTime.UtcNow - startUtc.Value;
        var seconds = (long)Math.Floor(elapsed.TotalSeconds);
        return Math.Max(0, seconds);
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
            TicketSecondsBooked = reader["ticket_seconds_booked"] == DBNull.Value ? Convert.ToInt64(reader["ticket_minutes_booked"]) * 60L : Convert.ToInt64(reader["ticket_seconds_booked"]),
            IsPinned = Convert.ToInt32(reader["is_pinned"]) != 0,
            IsZnunyAssigned = Convert.ToInt32(reader["is_znuny_assigned"]) != 0,
            CreatedUtc = ParseRequiredDateTime(reader["created_utc"].ToString()),
            UpdatedUtc = ParseRequiredDateTime(reader["updated_utc"].ToString()),
            TicketCreatedUtc = ParseNullableDateTime(reader["ticket_created_utc"]),
            TicketChangedUtc = ParseNullableDateTime(reader["ticket_changed_utc"]),
            LocalActivityUtc = ParseNullableDateTime(reader["local_activity_utc"])
                               ?? ParseRequiredDateTime(reader["created_utc"].ToString()),
            TicketState = reader["ticket_state"]?.ToString() ?? string.Empty ,
            TicketStateType = reader["ticket_state_type"]?.ToString() ?? string.Empty
        };
    }

    private static DateTime? ParseNullableDateTime(object value)
        => value == DBNull.Value || string.IsNullOrWhiteSpace(value.ToString())
            ? null
            : ParseRequiredDateTime(value.ToString()).ToUniversalTime();

    private static DateTime ParseRequiredDateTime(string? value)
    {
        if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var parsed))
            return parsed;

        if (DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.AssumeLocal, out parsed))
            return parsed;

        throw new FormatException($"Ungültiger Datumswert in der Datenbank: '{value ?? "<null>"}'");
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
        if (task.TicketSecondsBooked <= 0 && task.TicketMinutesBooked > 0)
            task.TicketSecondsBooked = task.TicketMinutesBooked * 60L;
        task.TicketMinutesBooked = (int)(Math.Max(0, task.TicketSecondsBooked) / 60);

        cmd.Parameters.AddWithValue("$ticket", task.TicketMinutesBooked);
        cmd.Parameters.AddWithValue("$ticketSeconds", task.TicketSecondsBooked);
        cmd.Parameters.AddWithValue("$pinned", task.IsPinned ? 1 : 0);
        cmd.Parameters.AddWithValue("$znunyAssigned", task.IsZnunyAssigned ? 1 : 0);
        cmd.Parameters.AddWithValue("$created", task.CreatedUtc.ToString("O"));
        cmd.Parameters.AddWithValue("$updated", task.UpdatedUtc.ToString("O"));
        cmd.Parameters.AddWithValue("$ticketCreated", task.TicketCreatedUtc?.ToString("O") ?? (object)DBNull.Value);
        cmd.Parameters.AddWithValue("$ticketChanged", task.TicketChangedUtc?.ToString("O") ?? (object)DBNull.Value);
        cmd.Parameters.AddWithValue("$localActivity", task.LocalActivityUtc.ToString("O"));
        cmd.Parameters.AddWithValue("$ticketState", task.TicketState);
        cmd.Parameters.AddWithValue("$ticketStateType", task.TicketStateType);
    }
}
