using System.Globalization;
using Microsoft.Data.Sqlite;

namespace TaskTool.Services;

public sealed record CandidateScanState(string TicketId, DateTime LastSeenUtc, DateTime LastEvaluatedUtc, bool Matched, DateTime? RemoteChangedUtc);

public sealed class TicketCandidateScanStateService
{
    private readonly DatabaseService _database;
    public TicketCandidateScanStateService(DatabaseService database) => _database = database;

    public IReadOnlyDictionary<string, CandidateScanState> Load()
    {
        var result = new Dictionary<string, CandidateScanState>(StringComparer.OrdinalIgnoreCase);
        using var connection = new SqliteConnection(_database.ConnectionString); connection.Open();
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT ticket_id,last_seen_utc,last_evaluated_utc,matched,remote_changed_utc FROM znuny_candidate_scan_state";
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            var id = reader.GetString(0);
            result[id] = new CandidateScanState(id, Parse(reader.GetString(1)) ?? DateTime.MinValue,
                Parse(reader.GetString(2)) ?? DateTime.MinValue, reader.GetInt64(3) != 0,
                reader.IsDBNull(4) ? null : Parse(reader.GetString(4)));
        }
        return result;
    }

    public void Replace(IEnumerable<CandidateScanState> states)
    {
        using var connection = new SqliteConnection(_database.ConnectionString); connection.Open();
        using var transaction = connection.BeginTransaction();
        using (var delete = connection.CreateCommand()) { delete.Transaction = transaction; delete.CommandText = "DELETE FROM znuny_candidate_scan_state"; delete.ExecuteNonQuery(); }
        foreach (var state in states)
        {
            using var insert = connection.CreateCommand(); insert.Transaction = transaction;
            insert.CommandText = @"INSERT INTO znuny_candidate_scan_state(ticket_id,last_seen_utc,last_evaluated_utc,matched,remote_changed_utc)
VALUES($id,$seen,$evaluated,$matched,$changed)";
            insert.Parameters.AddWithValue("$id", state.TicketId);
            insert.Parameters.AddWithValue("$seen", state.LastSeenUtc.ToUniversalTime().ToString("O"));
            insert.Parameters.AddWithValue("$evaluated", state.LastEvaluatedUtc.ToUniversalTime().ToString("O"));
            insert.Parameters.AddWithValue("$matched", state.Matched ? 1 : 0);
            insert.Parameters.AddWithValue("$changed", (object?)state.RemoteChangedUtc?.ToUniversalTime().ToString("O") ?? DBNull.Value);
            insert.ExecuteNonQuery();
        }
        transaction.Commit();
    }

    public void Remove(string ticketId)
    {
        using var connection = new SqliteConnection(_database.ConnectionString); connection.Open();
        using var command = connection.CreateCommand(); command.CommandText = "DELETE FROM znuny_candidate_scan_state WHERE ticket_id=$id";
        command.Parameters.AddWithValue("$id", ticketId); command.ExecuteNonQuery();
    }

    private static DateTime? Parse(string? value) => DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var parsed) ? parsed.ToUniversalTime() : null;
}
