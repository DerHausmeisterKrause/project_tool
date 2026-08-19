using Microsoft.Data.Sqlite;
using System.Globalization;
using TaskTool.Models;

namespace TaskTool.Services;

public class WorkDayService
{
    private readonly DatabaseService _db;
    private readonly LoggerService _logger;

    public WorkDayService(DatabaseService db, LoggerService logger)
    {
        _db = db;
        _logger = logger;
    }

    public WorkDayRecord GetOrCreateToday() => GetOrCreateDay(DateTime.Today.ToString("yyyy-MM-dd"));

    public WorkDayRecord GetOrCreateDay(string day)
    {
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var select = conn.CreateCommand();
        select.CommandText = "SELECT day,come_local,go_local,day_type,is_br,is_ho,homeoffice_outlook_entry_id FROM work_days WHERE day=$d";
        select.Parameters.AddWithValue("$d", day);
        using var r = select.ExecuteReader();
        if (r.Read()) return MapWorkDay(r, day);

        using var ins = conn.CreateCommand();
        ins.CommandText = "INSERT INTO work_days(day,day_type,is_br,is_ho) VALUES ($d,'Normal',0,0)";
        ins.Parameters.AddWithValue("$d", day);
        ins.ExecuteNonQuery();
        return new WorkDayRecord { Day = day, DayType = "Normal" };
    }

    public List<WorkDayRecord> GetWorkDaysInRange(DateTime from, DateTime to)
    {
        var list = new List<WorkDayRecord>();
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT day,come_local,go_local,day_type,is_br,is_ho,homeoffice_outlook_entry_id FROM work_days WHERE day>= $f AND day<= $t ORDER BY day";
        cmd.Parameters.AddWithValue("$f", from.ToString("yyyy-MM-dd"));
        cmd.Parameters.AddWithValue("$t", to.ToString("yyyy-MM-dd"));
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            var day = r["day"].ToString() ?? string.Empty;
            list.Add(MapWorkDay(r, day));
        }
        return list;
    }

    public void ReplaceSyncedOutlookMarkers(
        DateTime fromInclusive,
        DateTime toExclusive,
        IReadOnlyList<OutlookCalendarEvent> events,
        bool interpretAllDayMarkers)
    {
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var transaction = conn.BeginTransaction();
        for (var day = fromInclusive.Date; day < toExclusive.Date; day = day.AddDays(1))
        {
            var markers = CalendarMarkerResolver.ResolveOutlookMarkers(day, events, interpretAllDayMarkers, _logger);
            using var cmd = conn.CreateCommand();
            cmd.Transaction = transaction;
            cmd.CommandText = @"INSERT INTO calendar_marker_sync(day,outlook_day_type,outlook_is_ho,updated_utc)
VALUES ($day,$dayType,$isHo,$updated)
ON CONFLICT(day) DO UPDATE SET
    outlook_day_type=excluded.outlook_day_type,
    outlook_is_ho=excluded.outlook_is_ho,
    updated_utc=excluded.updated_utc";
            cmd.Parameters.AddWithValue("$day", markers.Day);
            cmd.Parameters.AddWithValue("$dayType", markers.OutlookDayType);
            cmd.Parameters.AddWithValue("$isHo", markers.OutlookIsHo ? 1 : 0);
            cmd.Parameters.AddWithValue("$updated", DateTime.UtcNow.ToString("O"));
            cmd.ExecuteNonQuery();
        }
        transaction.Commit();
    }

    public IReadOnlyList<MonthlyWorkDayStats> GetMonthlyMarkerStatistics()
    {
        var monthly = new Dictionary<DateTime, (int HomeOffice, int Vacation, int Am)>();
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = @"WITH days AS (
    SELECT day FROM work_days
    UNION
    SELECT day FROM calendar_marker_sync
    WHERE outlook_day_type IN ('UL', 'AM') OR outlook_is_ho = 1
)
SELECT d.day,
       COALESCE(w.day_type, 'Normal') AS local_day_type,
       COALESCE(w.is_ho, 0) AS local_is_ho,
       COALESCE(c.outlook_day_type, 'Normal') AS outlook_day_type,
       COALESCE(c.outlook_is_ho, 0) AS outlook_is_ho
FROM days d
LEFT JOIN work_days w ON w.day = d.day
LEFT JOIN calendar_marker_sync c ON c.day = d.day
ORDER BY d.day";
        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            var dayText = reader["day"]?.ToString() ?? string.Empty;
            if (!DateTime.TryParseExact(dayText, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var day))
                continue;

            var local = new WorkDayRecord
            {
                Day = dayText,
                DayType = reader["local_day_type"]?.ToString() ?? "Normal",
                IsHo = Convert.ToInt32(reader["local_is_ho"]) == 1
            };
            var outlook = new SyncedCalendarMarkers(
                dayText,
                reader["outlook_day_type"]?.ToString() ?? "Normal",
                Convert.ToInt32(reader["outlook_is_ho"]) == 1);
            var effective = CalendarMarkerResolver.ResolveEffectiveMarkers(local, outlook);
            var month = new DateTime(day.Year, day.Month, 1);
            var current = monthly.GetValueOrDefault(month);
            monthly[month] = (
                current.HomeOffice + (effective.IsHo ? 1 : 0),
                current.Vacation + (effective.DayType == "UL" ? 1 : 0),
                current.Am + (effective.DayType == "AM" ? 1 : 0));
        }

        return monthly
            .OrderByDescending(pair => pair.Key)
            .Select(pair => new MonthlyWorkDayStats(pair.Key, pair.Value.HomeOffice, pair.Value.Vacation, pair.Value.Am))
            .ToList();
    }

    public List<BreakRecord> GetBreaks(string day)
    {
        var list = new List<BreakRecord>();
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT * FROM breaks WHERE day=$d ORDER BY start_local";
        cmd.Parameters.AddWithValue("$d", day);
        using var r = cmd.ExecuteReader();
        while (r.Read())
        {
            list.Add(new BreakRecord
            {
                Id = Convert.ToInt64(r["id"]),
                Day = day,
                StartLocal = ParseRequiredDateTime(r["start_local"].ToString()),
                EndLocal = DateTime.TryParse(r["end_local"]?.ToString(), out var e) ? e : null,
                Note = r["note"]?.ToString() ?? string.Empty
            });
        }
        return list;
    }

    public void SetCome(DateTime time)
    {
        var record = GetOrCreateToday();
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE work_days SET come_local=$t WHERE day=$d";
        cmd.Parameters.AddWithValue("$t", time.ToString("s"));
        cmd.Parameters.AddWithValue("$d", record.Day);
        cmd.ExecuteNonQuery();
    }

    public void SetGo(DateTime time)
    {
        var record = GetOrCreateToday();
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE work_days SET go_local=$t WHERE day=$d";
        cmd.Parameters.AddWithValue("$t", time.ToString("s"));
        cmd.Parameters.AddWithValue("$d", record.Day);
        cmd.ExecuteNonQuery();
    }

    public void SetDayMarkers(string day, string dayType, bool isBr, bool isHo)
    {
        GetOrCreateDay(day);
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE work_days SET day_type=$t,is_br=$br,is_ho=$ho WHERE day=$d";
        cmd.Parameters.AddWithValue("$t", dayType);
        cmd.Parameters.AddWithValue("$br", isBr ? 1 : 0);
        cmd.Parameters.AddWithValue("$ho", isHo ? 1 : 0);
        cmd.Parameters.AddWithValue("$d", day);
        cmd.ExecuteNonQuery();
    }

    public void SetHomeOfficeState(string day, bool isHo, string entryId)
    {
        GetOrCreateDay(day);
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE work_days SET is_ho=$ho,homeoffice_outlook_entry_id=$entry WHERE day=$d";
        cmd.Parameters.AddWithValue("$ho", isHo ? 1 : 0);
        cmd.Parameters.AddWithValue("$entry", string.IsNullOrWhiteSpace(entryId) ? DBNull.Value : entryId);
        cmd.Parameters.AddWithValue("$d", day);
        if (cmd.ExecuteNonQuery() != 1)
            throw new InvalidOperationException("Homeoffice-Marker konnte nicht gespeichert werden.");
    }

    public void StartBreak(string day)
    {
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "INSERT INTO breaks(day,start_local,note) VALUES ($d,$s,'pause')";
        cmd.Parameters.AddWithValue("$d", day);
        cmd.Parameters.AddWithValue("$s", DateTime.Now.ToString("s"));
        cmd.ExecuteNonQuery();
    }

    public void EndBreak(string day)
    {
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "UPDATE breaks SET end_local=$e WHERE id=(SELECT id FROM breaks WHERE day=$d AND end_local IS NULL ORDER BY id DESC LIMIT 1)";
        cmd.Parameters.AddWithValue("$d", day);
        cmd.Parameters.AddWithValue("$e", DateTime.Now.ToString("s"));
        cmd.ExecuteNonQuery();
    }

    public void SaveManualDay(string day, DateTime? come, DateTime? go, IEnumerable<BreakRecord> breaks)
    {
        using var conn = new SqliteConnection(_db.ConnectionString);
        conn.Open();
        using var tx = conn.BeginTransaction();

        using (var upsert = conn.CreateCommand())
        {
            upsert.Transaction = tx;
            upsert.CommandText = @"INSERT INTO work_days(day,come_local,go_local,day_type,is_br,is_ho) VALUES ($d,$c,$g,'Normal',0,0)
ON CONFLICT(day) DO UPDATE SET come_local=$c, go_local=$g";
            upsert.Parameters.AddWithValue("$d", day);
            upsert.Parameters.AddWithValue("$c", come?.ToString("s") ?? (object)DBNull.Value);
            upsert.Parameters.AddWithValue("$g", go?.ToString("s") ?? (object)DBNull.Value);
            upsert.ExecuteNonQuery();
        }

        using (var del = conn.CreateCommand())
        {
            del.Transaction = tx;
            del.CommandText = "DELETE FROM breaks WHERE day=$d";
            del.Parameters.AddWithValue("$d", day);
            del.ExecuteNonQuery();
        }

        foreach (var br in breaks)
        {
            using var ins = conn.CreateCommand();
            ins.Transaction = tx;
            ins.CommandText = "INSERT INTO breaks(day,start_local,end_local,note) VALUES ($d,$s,$e,$n)";
            ins.Parameters.AddWithValue("$d", day);
            ins.Parameters.AddWithValue("$s", br.StartLocal.ToString("s"));
            ins.Parameters.AddWithValue("$e", br.EndLocal?.ToString("s") ?? (object)DBNull.Value);
            ins.Parameters.AddWithValue("$n", string.IsNullOrWhiteSpace(br.Note) ? "pause" : br.Note);
            ins.ExecuteNonQuery();
        }

        tx.Commit();
    }

    private static DateTime ParseRequiredDateTime(string? value)
    {
        if (DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var parsed))
            return parsed;

        if (DateTime.TryParse(value, CultureInfo.CurrentCulture, DateTimeStyles.AssumeLocal, out parsed))
            return parsed;

        throw new FormatException($"Ungültiger Datumswert in der Datenbank: '{value ?? "<null>"}'");
    }

    private static WorkDayRecord MapWorkDay(SqliteDataReader r, string day) => new()
    {
        Day = day,
        ComeLocal = DateTime.TryParse(r["come_local"]?.ToString(), out var c) ? c : null,
        GoLocal = DateTime.TryParse(r["go_local"]?.ToString(), out var g) ? g : null,
        DayType = r["day_type"]?.ToString() ?? "Normal",
        IsBr = Convert.ToInt32(r["is_br"]) == 1,
        IsHo = Convert.ToInt32(r["is_ho"]) == 1,
        HomeOfficeOutlookEntryId = r["homeoffice_outlook_entry_id"]?.ToString() ?? string.Empty
    };
}
