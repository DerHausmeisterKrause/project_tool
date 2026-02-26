using Microsoft.Data.Sqlite;
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
        select.CommandText = "SELECT day,come_local,go_local,day_type,is_br,is_ho FROM work_days WHERE day=$d";
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
        cmd.CommandText = "SELECT day,come_local,go_local,day_type,is_br,is_ho FROM work_days WHERE day>= $f AND day<= $t ORDER BY day";
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
                StartLocal = DateTime.Parse(r["start_local"].ToString()!),
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

    private static WorkDayRecord MapWorkDay(SqliteDataReader r, string day) => new()
    {
        Day = day,
        ComeLocal = DateTime.TryParse(r["come_local"]?.ToString(), out var c) ? c : null,
        GoLocal = DateTime.TryParse(r["go_local"]?.ToString(), out var g) ? g : null,
        DayType = r["day_type"]?.ToString() ?? "Normal",
        IsBr = Convert.ToInt32(r["is_br"]) == 1,
        IsHo = Convert.ToInt32(r["is_ho"]) == 1
    };
}
