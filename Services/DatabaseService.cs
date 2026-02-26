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

            if (currentVersion < 5)
            {
                SetVersion(conn, 5);
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
    go_local TEXT NULL
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
