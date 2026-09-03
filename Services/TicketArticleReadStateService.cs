using System.Globalization;
using System.Text.Json;
using Microsoft.Data.Sqlite;
using TaskTool.Models;

namespace TaskTool.Services;

/// <summary>Client-local read state. This service never communicates with Znuny.</summary>
public sealed class TicketArticleReadStateService
{
    private readonly DatabaseService _database;

    public TicketArticleReadStateService(DatabaseService database)
    {
        _database = database;
        InitializeKnownArticlesAsRead();
    }

    public void InitializeKnownArticlesAsRead()
    {
        using var connection = Open();
        using var transaction = connection.BeginTransaction();
        var ticketIds = new List<string>();
        using (var tickets = connection.CreateCommand())
        {
            tickets.Transaction = transaction;
            tickets.CommandText = "SELECT ticket_id FROM znuny_ticket_detail_cache";
            using var reader = tickets.ExecuteReader();
            while (reader.Read()) ticketIds.Add(reader.GetString(0));
        }
        foreach (var ticketId in ticketIds)
        {
            if (HasBaseline(connection, transaction, ticketId)) continue;
            var articles = LoadCachedArticles(connection, transaction, ticketId);
            CreateBaseline(connection, transaction, ticketId, articles, markKnownRead: true);
        }
        transaction.Commit();
    }

    /// <returns>True when the unread set changed.</returns>
    public bool ReconcileFetchedArticles(string ticketId, IReadOnlyCollection<TicketArticleItem> fetchedArticles)
    {
        if (string.IsNullOrWhiteSpace(ticketId)) return false;
        var articles = fetchedArticles.Where(a => !string.IsNullOrWhiteSpace(a.ArticleId)).ToList();
        using var connection = Open();
        using var transaction = connection.BeginTransaction();
        if (!HasBaseline(connection, transaction, ticketId))
        {
            CreateBaseline(connection, transaction, ticketId, articles, markKnownRead: true);
            transaction.Commit();
            return false;
        }

        var (watermarkCreated, watermarkId) = LoadWatermark(connection, transaction, ticketId);
        if (!watermarkCreated.HasValue && string.IsNullOrWhiteSpace(watermarkId))
        {
            foreach (var article in articles)
                InsertState(connection, transaction, ticketId, article.ArticleId, DateTime.UtcNow);
            var initialNewest = FindNewest(articles);
            if (initialNewest != null)
                UpdateWatermark(connection, transaction, ticketId, initialNewest.CreatedLocal, initialNewest.ArticleId);
            transaction.Commit();
            return false;
        }
        var changed = false;
        foreach (var article in articles)
        {
            if (Exists(connection, transaction, ticketId, article.ArticleId)) continue;
            var isNew = IsReliablyAfter(article.CreatedLocal, article.ArticleId, watermarkCreated, watermarkId);
            InsertState(connection, transaction, ticketId, article.ArticleId, isNew ? null : DateTime.UtcNow);
            changed |= isNew;
        }
        var newest = FindNewest(articles, watermarkCreated, watermarkId);
        if (newest != null && IsReliablyAfter(newest.CreatedLocal, newest.ArticleId, watermarkCreated, watermarkId))
            UpdateWatermark(connection, transaction, ticketId, newest.CreatedLocal, newest.ArticleId);
        transaction.Commit();
        return changed;
    }

    public IReadOnlySet<string> GetUnreadArticleIds(string ticketId)
    {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using var connection = Open(); using var command = connection.CreateCommand();
        command.CommandText = "SELECT article_id FROM znuny_ticket_article_read_state WHERE ticket_id=$id AND read_utc IS NULL";
        command.Parameters.AddWithValue("$id", ticketId);
        using var reader = command.ExecuteReader(); while (reader.Read()) result.Add(reader.GetString(0));
        return result;
    }

    public int GetUnreadCount(string ticketId)
    {
        using var connection = Open(); using var command = connection.CreateCommand();
        command.CommandText = "SELECT COUNT(*) FROM znuny_ticket_article_read_state WHERE ticket_id=$id AND read_utc IS NULL";
        command.Parameters.AddWithValue("$id", ticketId); return Convert.ToInt32(command.ExecuteScalar());
    }

    public bool HasUnread(string ticketId) => GetUnreadCount(ticketId) > 0;
    public bool IsUnread(string ticketId, string articleId) => GetUnreadArticleIds(ticketId).Contains(articleId);

    public bool MarkRead(string ticketId, string articleId)
    {
        if (string.IsNullOrWhiteSpace(ticketId) || string.IsNullOrWhiteSpace(articleId)) return false;
        using var connection = Open(); using var transaction = connection.BeginTransaction();
        using var command = connection.CreateCommand(); command.Transaction = transaction;
        command.CommandText = @"INSERT INTO znuny_ticket_article_read_state(ticket_id,article_id,first_seen_utc,read_utc)
VALUES($ticket,$article,$now,$now) ON CONFLICT(ticket_id,article_id) DO UPDATE SET read_utc=COALESCE(read_utc,excluded.read_utc)";
        command.Parameters.AddWithValue("$ticket", ticketId); command.Parameters.AddWithValue("$article", articleId);
        command.Parameters.AddWithValue("$now", DateTime.UtcNow.ToString("O"));
        var changed = command.ExecuteNonQuery() > 0; transaction.Commit(); return changed;
    }

    private static bool IsReliablyAfter(DateTime? created, string id, DateTime? watermarkCreated, string watermarkId)
    {
        if (created.HasValue && watermarkCreated.HasValue)
        {
            var compare = DateTime.Compare(created.Value.ToUniversalTime(), watermarkCreated.Value.ToUniversalTime());
            if (compare != 0) return compare > 0;
            return TryCompareNumericIds(id, watermarkId, out var idCompare) && idCompare > 0;
        }
        return TryCompareNumericIds(id, watermarkId, out var fallbackCompare) && fallbackCompare > 0;
    }

    private static bool TryCompareNumericIds(string left, string right, out int comparison)
    {
        if (long.TryParse(left, NumberStyles.None, CultureInfo.InvariantCulture, out var l)
            && long.TryParse(right, NumberStyles.None, CultureInfo.InvariantCulture, out var r))
        {
            comparison = l.CompareTo(r);
            return true;
        }
        comparison = 0;
        return false;
    }

    private static TicketArticleItem? FindNewest(IEnumerable<TicketArticleItem> articles, DateTime? created, string id)
        => FindNewest(articles.Where(a => IsReliablyAfter(a.CreatedLocal, a.ArticleId, created, id)));

    private static TicketArticleItem? FindNewest(IEnumerable<TicketArticleItem> articles)
    {
        TicketArticleItem? newest = null;
        foreach (var article in articles)
        {
            if (newest == null || IsReliablyAfter(article.CreatedLocal, article.ArticleId, newest.CreatedLocal, newest.ArticleId))
                newest = article;
        }
        return newest;
    }

    private static void CreateBaseline(SqliteConnection connection, SqliteTransaction transaction, string ticketId,
        IReadOnlyCollection<TicketArticleItem> articles, bool markKnownRead)
    {
        var newest = FindNewest(articles);
        using (var baseline = connection.CreateCommand()) { baseline.Transaction = transaction; baseline.CommandText = @"INSERT OR IGNORE INTO znuny_ticket_article_read_baseline(ticket_id,initialized_utc,newest_seen_created_utc,newest_seen_article_id) VALUES($id,$now,$created,$article)"; baseline.Parameters.AddWithValue("$id", ticketId); baseline.Parameters.AddWithValue("$now", DateTime.UtcNow.ToString("O")); baseline.Parameters.AddWithValue("$created", (object?)newest?.CreatedLocal?.ToUniversalTime().ToString("O") ?? DBNull.Value); baseline.Parameters.AddWithValue("$article", (object?)newest?.ArticleId ?? DBNull.Value); baseline.ExecuteNonQuery(); }
        if (markKnownRead) foreach (var article in articles.Where(a => !string.IsNullOrWhiteSpace(a.ArticleId))) InsertState(connection, transaction, ticketId, article.ArticleId, DateTime.UtcNow);
    }
    private static void InsertState(SqliteConnection connection, SqliteTransaction transaction, string ticketId, string articleId, DateTime? readUtc) { using var command = connection.CreateCommand(); command.Transaction = transaction; command.CommandText = "INSERT OR IGNORE INTO znuny_ticket_article_read_state(ticket_id,article_id,first_seen_utc,read_utc) VALUES($ticket,$article,$now,$read)"; command.Parameters.AddWithValue("$ticket", ticketId); command.Parameters.AddWithValue("$article", articleId); command.Parameters.AddWithValue("$now", DateTime.UtcNow.ToString("O")); command.Parameters.AddWithValue("$read", (object?)readUtc?.ToString("O") ?? DBNull.Value); command.ExecuteNonQuery(); }
    private static bool HasBaseline(SqliteConnection c, SqliteTransaction t, string id) { using var q=c.CreateCommand(); q.Transaction=t; q.CommandText="SELECT 1 FROM znuny_ticket_article_read_baseline WHERE ticket_id=$id"; q.Parameters.AddWithValue("$id",id); return q.ExecuteScalar()!=null; }
    private static bool Exists(SqliteConnection c, SqliteTransaction t, string ticket, string article) { using var q=c.CreateCommand(); q.Transaction=t; q.CommandText="SELECT 1 FROM znuny_ticket_article_read_state WHERE ticket_id=$ticket AND article_id=$article"; q.Parameters.AddWithValue("$ticket",ticket); q.Parameters.AddWithValue("$article",article); return q.ExecuteScalar()!=null; }
    private static (DateTime?,string) LoadWatermark(SqliteConnection c, SqliteTransaction t,string id) { using var q=c.CreateCommand(); q.Transaction=t;q.CommandText="SELECT newest_seen_created_utc,newest_seen_article_id FROM znuny_ticket_article_read_baseline WHERE ticket_id=$id";q.Parameters.AddWithValue("$id",id);using var r=q.ExecuteReader();r.Read();return (r.IsDBNull(0)?null:DateTime.Parse(r.GetString(0),null,DateTimeStyles.RoundtripKind),r.IsDBNull(1)?"":r.GetString(1)); }
    private static void UpdateWatermark(SqliteConnection c,SqliteTransaction t,string id,DateTime? created,string article){using var q=c.CreateCommand();q.Transaction=t;q.CommandText="UPDATE znuny_ticket_article_read_baseline SET newest_seen_created_utc=$created,newest_seen_article_id=$article WHERE ticket_id=$id";q.Parameters.AddWithValue("$id",id);q.Parameters.AddWithValue("$created",(object?)created?.ToUniversalTime().ToString("O")??DBNull.Value);q.Parameters.AddWithValue("$article",article);q.ExecuteNonQuery();}
    private static List<TicketArticleItem> LoadCachedArticles(SqliteConnection c,SqliteTransaction t,string id){var result=new List<TicketArticleItem>();using var q=c.CreateCommand();q.Transaction=t;q.CommandText="SELECT payload_json FROM znuny_ticket_article_cache WHERE ticket_id=$id";q.Parameters.AddWithValue("$id",id);using var r=q.ExecuteReader();while(r.Read()){try{var a=JsonSerializer.Deserialize<TicketArticleItem>(r.GetString(0));if(a!=null&&!string.IsNullOrWhiteSpace(a.ArticleId))result.Add(a);}catch(JsonException){}}return result;}
    private SqliteConnection Open(){var c=new SqliteConnection(_database.ConnectionString);c.Open();return c;}
}
