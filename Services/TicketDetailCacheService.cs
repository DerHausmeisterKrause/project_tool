using System.Globalization;
using System.Text.Json;
using Microsoft.Data.Sqlite;
using TaskTool.Models;

namespace TaskTool.Services;

public sealed class TicketDetailCacheService
{
    private readonly DatabaseService _database;
    public TicketDetailCacheService(DatabaseService database) => _database = database;

    public DateTime? GetRemoteChangedUtc(string ticketId)
    {
        using var connection = Open();
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT remote_changed_utc FROM znuny_ticket_detail_cache WHERE ticket_id=$id";
        command.Parameters.AddWithValue("$id", ticketId);
        return Parse(command.ExecuteScalar()?.ToString());
    }

    public TicketBookingContext? Load(string ticketId)
    {
        using var connection = Open();
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT * FROM znuny_ticket_detail_cache WHERE ticket_id=$id";
        command.Parameters.AddWithValue("$id", ticketId);
        using var reader = command.ExecuteReader();
        if (!reader.Read()) return null;
        var number = reader["ticket_number"].ToString() ?? "";
        var title = reader["title"].ToString() ?? "";
        var cost = reader["cost_center_value"].ToString() ?? "";
        var order = reader["order_value"].ToString() ?? "";
        var recipient = reader["reply_recipient"].ToString() ?? "";
        var replyId = reader["reply_source_article_id"].ToString() ?? "";
        reader.Close();
        var articles = LoadArticles(connection, ticketId);
        return new(ticketId, number, cost, order, Array.Empty<TicketFieldOption>(),
            Array.Empty<TicketFieldOption>(), "", articles,
            articles.FirstOrDefault(a => a.ArticleId == replyId), recipient, title);
    }

    public void Store(TicketBookingContext context, string state, DateTime? remoteChangedUtc)
    {
        using var connection = Open();
        using var transaction = connection.BeginTransaction();
        using (var command = connection.CreateCommand())
        {
            command.Transaction = transaction;
            command.CommandText = @"INSERT INTO znuny_ticket_detail_cache
(ticket_id,ticket_number,title,state,remote_changed_utc,last_fetched_utc,cost_center_value,order_value,reply_recipient,reply_source_article_id)
VALUES($id,$number,$title,$state,$changed,$fetched,$cost,$order,$recipient,$reply)
ON CONFLICT(ticket_id) DO UPDATE SET ticket_number=excluded.ticket_number,title=excluded.title,state=excluded.state,
remote_changed_utc=excluded.remote_changed_utc,last_fetched_utc=excluded.last_fetched_utc,cost_center_value=excluded.cost_center_value,
order_value=excluded.order_value,reply_recipient=excluded.reply_recipient,reply_source_article_id=excluded.reply_source_article_id";
            command.Parameters.AddWithValue("$id", context.TicketId);
            command.Parameters.AddWithValue("$number", context.TicketNumber);
            command.Parameters.AddWithValue("$title", context.TicketTitle);
            command.Parameters.AddWithValue("$state", state);
            command.Parameters.AddWithValue("$changed", (object?)remoteChangedUtc?.ToUniversalTime().ToString("O") ?? DBNull.Value);
            command.Parameters.AddWithValue("$fetched", DateTime.UtcNow.ToString("O"));
            command.Parameters.AddWithValue("$cost", context.CostCenterValue);
            command.Parameters.AddWithValue("$order", context.OrderValue);
            command.Parameters.AddWithValue("$recipient", context.ReplyRecipient);
            command.Parameters.AddWithValue("$reply", context.ReplySourceArticle?.ArticleId ?? "");
            command.ExecuteNonQuery();
        }
        using (var delete = connection.CreateCommand()) { delete.Transaction = transaction; delete.CommandText = "DELETE FROM znuny_ticket_article_cache WHERE ticket_id=$id"; delete.Parameters.AddWithValue("$id", context.TicketId); delete.ExecuteNonQuery(); }
        for (var index = 0; index < context.Articles.Count; index++)
        {
            using var insert = connection.CreateCommand();
            insert.Transaction = transaction;
            insert.CommandText = "INSERT INTO znuny_ticket_article_cache(ticket_id,article_id,ordinal,payload_json) VALUES($id,$article,$ordinal,$json)";
            insert.Parameters.AddWithValue("$id", context.TicketId); insert.Parameters.AddWithValue("$article", context.Articles[index].ArticleId);
            insert.Parameters.AddWithValue("$ordinal", index); insert.Parameters.AddWithValue("$json", JsonSerializer.Serialize(context.Articles[index])); insert.ExecuteNonQuery();
        }
        transaction.Commit();
    }

    private static IReadOnlyList<TicketArticleItem> LoadArticles(SqliteConnection connection, string ticketId)
    {
        var result = new List<TicketArticleItem>(); using var command = connection.CreateCommand();
        command.CommandText = "SELECT payload_json FROM znuny_ticket_article_cache WHERE ticket_id=$id ORDER BY ordinal"; command.Parameters.AddWithValue("$id", ticketId);
        using var reader = command.ExecuteReader(); while (reader.Read()) { var item = JsonSerializer.Deserialize<TicketArticleItem>(reader.GetString(0)); if (item != null) result.Add(item); }
        return result;
    }
    private SqliteConnection Open() { var connection = new SqliteConnection(_database.ConnectionString); connection.Open(); return connection; }
    private static DateTime? Parse(string? value) => DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var parsed) ? parsed.ToUniversalTime() : null;
}
