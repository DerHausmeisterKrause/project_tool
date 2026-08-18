using Microsoft.Data.Sqlite;

namespace TaskTool.Services;

public sealed record TicketAssignmentSnapshot(bool Initialized, HashSet<string> TicketIds);

public sealed class TicketAssignmentSnapshotService
{
    private readonly DatabaseService _database;

    public TicketAssignmentSnapshotService(DatabaseService database)
    {
        _database = database;
    }

    public TicketAssignmentSnapshot Load(string contextKey)
    {
        using var connection = new SqliteConnection(_database.ConnectionString);
        connection.Open();

        var initialized = false;
        using (var stateCommand = connection.CreateCommand())
        {
            stateCommand.CommandText = "SELECT initialized FROM ticket_assignment_sync_state WHERE context_key=$context";
            stateCommand.Parameters.AddWithValue("$context", contextKey);
            var value = stateCommand.ExecuteScalar();
            initialized = value != null && value != DBNull.Value && Convert.ToInt32(value) == 1;
        }

        var ticketIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using var snapshotCommand = connection.CreateCommand();
        snapshotCommand.CommandText = "SELECT ticket_id FROM ticket_assignment_snapshot WHERE context_key=$context";
        snapshotCommand.Parameters.AddWithValue("$context", contextKey);
        using var reader = snapshotCommand.ExecuteReader();
        while (reader.Read())
        {
            var ticketId = reader["ticket_id"]?.ToString();
            if (!string.IsNullOrWhiteSpace(ticketId)) ticketIds.Add(ticketId);
        }

        return new TicketAssignmentSnapshot(initialized, ticketIds);
    }

    public void Replace(string contextKey, IEnumerable<string> currentTicketIds)
    {
        var ticketIds = currentTicketIds
            .Where(ticketId => !string.IsNullOrWhiteSpace(ticketId))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        var updatedUtc = DateTime.UtcNow.ToString("O");

        using var connection = new SqliteConnection(_database.ConnectionString);
        connection.Open();
        using var transaction = connection.BeginTransaction();

        using (var deleteCommand = connection.CreateCommand())
        {
            deleteCommand.Transaction = transaction;
            deleteCommand.CommandText = "DELETE FROM ticket_assignment_snapshot WHERE context_key=$context";
            deleteCommand.Parameters.AddWithValue("$context", contextKey);
            deleteCommand.ExecuteNonQuery();
        }

        foreach (var ticketId in ticketIds)
        {
            using var insertCommand = connection.CreateCommand();
            insertCommand.Transaction = transaction;
            insertCommand.CommandText = "INSERT INTO ticket_assignment_snapshot(context_key,ticket_id,last_seen_utc) VALUES ($context,$ticket,$updated)";
            insertCommand.Parameters.AddWithValue("$context", contextKey);
            insertCommand.Parameters.AddWithValue("$ticket", ticketId);
            insertCommand.Parameters.AddWithValue("$updated", updatedUtc);
            insertCommand.ExecuteNonQuery();
        }

        using (var stateCommand = connection.CreateCommand())
        {
            stateCommand.Transaction = transaction;
            stateCommand.CommandText = @"INSERT INTO ticket_assignment_sync_state(context_key,initialized,updated_utc)
VALUES ($context,1,$updated)
ON CONFLICT(context_key) DO UPDATE SET initialized=1,updated_utc=excluded.updated_utc";
            stateCommand.Parameters.AddWithValue("$context", contextKey);
            stateCommand.Parameters.AddWithValue("$updated", updatedUtc);
            stateCommand.ExecuteNonQuery();
        }

        transaction.Commit();
    }
}
