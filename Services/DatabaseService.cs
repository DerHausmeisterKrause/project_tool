using Microsoft.Data.Sqlite;
using System.IO;

namespace TaskTool.Services;

public class DatabaseService
{
    private readonly LoggerService _logger;
    private readonly string _dbPath = Path.Combine(AppContext.BaseDirectory, "TaskTool.db");
    public string ConnectionString => new SqliteConnectionStringBuilder { DataSource = _dbPath }.ToString();

    public DatabaseService(LoggerService logger) => _logger = logger;

    public void Initialize()
    {
        try
        {
            using var conn = new SqliteConnection(ConnectionString);
            conn.Open();

            Exec(conn, @"
CREATE TABLE IF NOT EXISTS schema_version (
    version INTEGER NOT NULL
);");

            using (var check = conn.CreateCommand())
            {
                check.CommandText = "SELECT COUNT(*) FROM schema_version";
                var count = Convert.ToInt32(check.ExecuteScalar());
                if (count == 0)
                {
                    Exec(conn, "INSERT INTO schema_version(version) VALUES (1);");
                }
            }

            var currentVersion = GetVersion(conn);
            if (currentVersion < 1)
            {
                SetVersion(conn, 1);
                currentVersion = 1;
            }

            // Ensure the schema is fully present even on partially initialized databases.
            CreateBaseSchema(conn);
            MigrateToV2(conn);
            MigrateToV3(conn);
            MigrateToV4(conn);
            MigrateToV5(conn);
            MigrateToV6(conn);
            MigrateToV7(conn);
            MigrateToV8(conn);
            MigrateToV9(conn);
            MigrateToV10(conn);
            MigrateToV11(conn);

            if (currentVersion < 11)
            {
                SetVersion(conn, 11);
            }
        }
        catch (Exception ex)
        {
            _logger.Error($"DB init failed: {ex.Message}");
        }
    }

    private static void CreateBaseSchema(SqliteConnection conn)
    {
        Exec(conn, @"
CREATE TABLE IF NOT EXISTS tasks (
    id TEXT PRIMARY KEY,
    title TEXT NOT NULL,
    description TEXT,
    ticket_url TEXT,
    start_local TEXT NULL,
    end_local TEXT NULL,
    status TEXT NOT NULL,
    priority INTEGER NULL,
    tags TEXT,
    outlook_entry_id TEXT,
    ticket_minutes_booked INTEGER NOT NULL DEFAULT 0,
    ticket_seconds_booked INTEGER NOT NULL DEFAULT 0,
    is_pinned INTEGER NOT NULL DEFAULT 0,
    created_utc TEXT NOT NULL,
    updated_utc TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS time_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    task_id TEXT NOT NULL,
    start_utc TEXT NOT NULL,
    end_utc TEXT NULL,
    note TEXT
);

CREATE TABLE IF NOT EXISTS work_days (
    day TEXT PRIMARY KEY,
    come_local TEXT NULL,
    go_local TEXT NULL,
    homeoffice_outlook_entry_id TEXT NULL
);

CREATE TABLE IF NOT EXISTS breaks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    day TEXT NOT NULL,
    start_local TEXT NOT NULL,
    end_local TEXT NULL,
    note TEXT
); ");
    }

    private static void MigrateToV2(SqliteConnection conn)
    {
        EnsureColumn(conn, "work_days", "day_type", "TEXT NOT NULL DEFAULT 'Normal'");
        EnsureColumn(conn, "work_days", "is_br", "INTEGER NOT NULL DEFAULT 0");
        EnsureColumn(conn, "work_days", "is_ho", "INTEGER NOT NULL DEFAULT 0");
    }

    private static void MigrateToV3(SqliteConnection conn)
    {
        Exec(conn, @"CREATE TABLE IF NOT EXISTS task_segments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    task_id TEXT NOT NULL,
    start_local TEXT NOT NULL,
    end_local TEXT NOT NULL,
    planned_minutes INTEGER NOT NULL DEFAULT 0,
    note TEXT NOT NULL DEFAULT '',
    outlook_entry_id TEXT
);");
    }

    private static void MigrateToV4(SqliteConnection conn)
    {
        EnsureColumn(conn, "task_segments", "note", "TEXT NOT NULL DEFAULT ''");
    }

    private static void MigrateToV5(SqliteConnection conn)
    {
        EnsureColumn(conn, "tasks", "ticket_seconds_booked", "INTEGER NOT NULL DEFAULT 0");
        Exec(conn, "UPDATE tasks SET ticket_seconds_booked = CASE WHEN ticket_seconds_booked IS NULL OR ticket_seconds_booked <= 0 THEN ticket_minutes_booked * 60 ELSE ticket_seconds_booked END;");
    }

    private static void MigrateToV6(SqliteConnection conn)
    {
        Exec(conn, @"CREATE TABLE IF NOT EXISTS ticket_time_bookings (
    id TEXT PRIMARY KEY,
    task_id TEXT NOT NULL,
    ticket_id TEXT NOT NULL,
    ticket_number TEXT NOT NULL,
    booking_id TEXT NOT NULL UNIQUE,
    article_id TEXT NOT NULL DEFAULT '',
    minutes REAL NOT NULL,
    source_seconds INTEGER NOT NULL,
    booked_at_utc TEXT NOT NULL,
    short_description TEXT NOT NULL DEFAULT '',
    cost_center TEXT NOT NULL DEFAULT '',
    order_value TEXT NOT NULL DEFAULT '',
    status TEXT NOT NULL DEFAULT 'Pending'
);
CREATE INDEX IF NOT EXISTS idx_ticket_time_bookings_task_status
ON ticket_time_bookings(task_id, status, booked_at_utc DESC);
CREATE UNIQUE INDEX IF NOT EXISTS idx_ticket_time_bookings_one_pending_per_task
ON ticket_time_bookings(task_id) WHERE status = 'Pending';");
    }

    private static void MigrateToV7(SqliteConnection conn)
    {
        EnsureColumn(conn, "ticket_time_bookings", "booked_minutes", "REAL NOT NULL DEFAULT 0");
        Exec(conn, @"UPDATE ticket_time_bookings
SET booked_minutes = minutes
WHERE booked_minutes IS NULL OR booked_minutes <= 0;

CREATE TABLE IF NOT EXISTS ticket_time_booking_baselines (
    task_id TEXT PRIMARY KEY,
    baseline_seconds INTEGER NOT NULL,
    created_utc TEXT NOT NULL
);

INSERT OR IGNORE INTO ticket_time_booking_baselines(task_id, baseline_seconds, created_utc)
SELECT t.id,
       MAX(0, t.ticket_seconds_booked
           - COALESCE(SUM(b.source_seconds), 0)),
       datetime('now')
FROM tasks t
LEFT JOIN ticket_time_bookings b ON b.task_id = t.id
GROUP BY t.id;");
    }

    private static void MigrateToV8(SqliteConnection conn)
    {
        EnsureColumn(conn, "tasks", "is_pinned", "INTEGER NOT NULL DEFAULT 0");
    }

    private static void MigrateToV9(SqliteConnection conn)
    {
        EnsureColumn(conn, "work_days", "homeoffice_outlook_entry_id", "TEXT NULL");
    }

    private static void MigrateToV10(SqliteConnection conn)
    {
        Exec(conn, @"CREATE TABLE IF NOT EXISTS ticket_assignment_snapshot (
    context_key TEXT NOT NULL,
    ticket_id TEXT NOT NULL,
    last_seen_utc TEXT NOT NULL,
    PRIMARY KEY (context_key, ticket_id)
);
CREATE TABLE IF NOT EXISTS ticket_assignment_sync_state (
    context_key TEXT PRIMARY KEY,
    initialized INTEGER NOT NULL DEFAULT 0,
    updated_utc TEXT NOT NULL
);");
    }

    private static void MigrateToV11(SqliteConnection conn)
    {
        EnsureColumn(conn, "tasks", "is_znuny_assigned", "INTEGER NOT NULL DEFAULT 1");
    }

    private static void EnsureColumn(SqliteConnection conn, string table, string column, string definition)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = $"PRAGMA table_info({table})";
        using var reader = cmd.ExecuteReader();
        while (reader.Read())
        {
            if (string.Equals(reader["name"]?.ToString(), column, StringComparison.OrdinalIgnoreCase))
                return;
        }
        Exec(conn, $"ALTER TABLE {table} ADD COLUMN {column} {definition};");
    }

    private static int GetVersion(SqliteConnection conn)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT version FROM schema_version ORDER BY rowid DESC LIMIT 1";
        return Convert.ToInt32(cmd.ExecuteScalar());
    }

    private static void SetVersion(SqliteConnection conn, int version)
    {
        Exec(conn, "DELETE FROM schema_version;");
        using var cmd = conn.CreateCommand();
        cmd.CommandText = "INSERT INTO schema_version(version) VALUES ($v)";
        cmd.Parameters.AddWithValue("$v", version);
        cmd.ExecuteNonQuery();
    }

    private static void Exec(SqliteConnection conn, string sql)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        cmd.ExecuteNonQuery();
    }
}
