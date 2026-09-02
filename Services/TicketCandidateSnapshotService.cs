using System.Globalization;
using Microsoft.Data.Sqlite;
using TaskTool.Models;

namespace TaskTool.Services;

public sealed class TicketCandidateSnapshotService
{
    private readonly DatabaseService _database;
    public TicketCandidateSnapshotService(DatabaseService database) => _database = database;

    public IReadOnlyList<ZnunyCandidateTicket> Load() => LoadSnapshot("znuny_candidate_snapshot");
    public IReadOnlyList<ZnunyCandidateTicket> LoadPool() => LoadSnapshot("znuny_candidate_pool_snapshot");

    private IReadOnlyList<ZnunyCandidateTicket> LoadSnapshot(string table)
    {
        var result = new List<ZnunyCandidateTicket>();
        using var connection = new SqliteConnection(_database.ConnectionString);
        connection.Open();
        using var command = connection.CreateCommand();
        command.CommandText = $"SELECT * FROM {table} ORDER BY ticket_number DESC";
        using var reader = command.ExecuteReader();
        while (reader.Read())
        {
            result.Add(new ZnunyCandidateTicket
            {
                TicketId = reader["ticket_id"].ToString() ?? string.Empty,
                TicketNumber = reader["ticket_number"].ToString() ?? string.Empty,
                Title = reader["title"].ToString() ?? string.Empty,
                DescriptionPreview = reader["description_preview"].ToString() ?? string.Empty,
                Owner = reader["owner"].ToString() ?? string.Empty,
                Responsible = reader["responsible"].ToString() ?? string.Empty,
                State = reader["state"].ToString() ?? string.Empty,
                WebUrl = reader["web_url"].ToString() ?? string.Empty,
                MatchedKeyword = reader["matched_keyword"].ToString() ?? string.Empty,
                LastSyncedUtc = DateTime.TryParse(reader["last_synced_utc"].ToString(), CultureInfo.InvariantCulture,
                    DateTimeStyles.RoundtripKind, out var synced) ? synced : DateTime.MinValue
            });
        }
        return result;
    }

    public void Replace(IReadOnlyCollection<ZnunyCandidateTicket> tickets)
        => ReplaceSnapshot("znuny_candidate_snapshot", tickets);

    public void ReplacePool(IReadOnlyCollection<ZnunyCandidateTicket> tickets)
        => ReplaceSnapshot("znuny_candidate_pool_snapshot", tickets);

    private void ReplaceSnapshot(string table, IReadOnlyCollection<ZnunyCandidateTicket> tickets)
    {
        using var connection = new SqliteConnection(_database.ConnectionString);
        connection.Open();
        using var transaction = connection.BeginTransaction();
        using (var delete = connection.CreateCommand())
        {
            delete.Transaction = transaction;
            delete.CommandText = $"DELETE FROM {table}";
            delete.ExecuteNonQuery();
        }
        foreach (var ticket in tickets)
        {
            using var insert = connection.CreateCommand();
            insert.Transaction = transaction;
            insert.CommandText = $@"INSERT INTO {table}
(ticket_id,ticket_number,title,description_preview,owner,responsible,state,web_url,matched_keyword,last_synced_utc)
VALUES($id,$number,$title,$preview,$owner,$responsible,$state,$url,$keyword,$synced)";
            insert.Parameters.AddWithValue("$id", ticket.TicketId);
            insert.Parameters.AddWithValue("$number", ticket.TicketNumber);
            insert.Parameters.AddWithValue("$title", ticket.Title);
            insert.Parameters.AddWithValue("$preview", ticket.DescriptionPreview);
            insert.Parameters.AddWithValue("$owner", ticket.Owner);
            insert.Parameters.AddWithValue("$responsible", ticket.Responsible);
            insert.Parameters.AddWithValue("$state", ticket.State);
            insert.Parameters.AddWithValue("$url", ticket.WebUrl);
            insert.Parameters.AddWithValue("$keyword", ticket.MatchedKeyword);
            insert.Parameters.AddWithValue("$synced", ticket.LastSyncedUtc.ToUniversalTime().ToString("O"));
            insert.ExecuteNonQuery();
        }
        transaction.Commit();
    }
}
