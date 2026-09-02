using System.Globalization;
using System.Text.Json;
using Microsoft.Data.Sqlite;
using TaskTool.Models;

namespace TaskTool.Services;

public sealed class TicketDetailCacheService
{
    private readonly DatabaseService _database;
    public TicketDetailCacheService(DatabaseService database) => _database = database;

    public DateTime? GetRemoteChangedUtc(string ticketId) => LoadEntry(ticketId)?.RemoteChangedUtc;

    public TicketBookingContext? Load(string ticketId) => LoadEntry(ticketId)?.Context;

    public TicketDetailCacheEntry? LoadEntry(string ticketId)
    {
        using var connection = Open();
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT * FROM znuny_ticket_detail_cache WHERE ticket_id=$id";
        command.Parameters.AddWithValue("$id", ticketId);
        using var reader = command.ExecuteReader();
        if (!reader.Read()) return null;
        var context = new TicketBookingContext(ticketId, Text(reader, "ticket_number"), Text(reader, "cost_center_value"),
            Text(reader, "order_value"), Array.Empty<TicketFieldOption>(), Array.Empty<TicketFieldOption>(), "",
            Array.Empty<TicketArticleItem>(), null, Text(reader, "reply_recipient"), Text(reader, "title"));
        var state = Text(reader, "state");
        var changed = Parse(NullableText(reader, "remote_changed_utc"));
        var fetched = Parse(NullableText(reader, "last_fetched_utc")) ?? DateTime.MinValue;
        var metadataComplete = Bool(reader, "metadata_complete");
        var articlesComplete = Bool(reader, "articles_complete");
        var dynamicFieldsComplete = Bool(reader, "dynamic_fields_complete");
        var articleLimit = Int(reader, "fetched_article_limit");
        var replyId = Text(reader, "reply_source_article_id");
        reader.Close();
        var articles = LoadArticles(connection, ticketId);
        context = context with { Articles = articles, ReplySourceArticle = articles.FirstOrDefault(a => a.ArticleId == replyId) };
        return new TicketDetailCacheEntry(context, state, changed, fetched, metadataComplete,
            articlesComplete, dynamicFieldsComplete, articleLimit);
    }

    public void Store(TicketBookingContext context, string state, DateTime? remoteChangedUtc,
        TicketDetailFetchProfile? profile = null)
    {
        profile ??= new TicketDetailFetchProfile(true, false, false, 0);
        var existing = LoadEntry(context.TicketId);
        var sameRemoteVersion = existing != null && SameRemoteVersion(existing.RemoteChangedUtc, remoteChangedUtc);
        // A narrower Candidate read may refresh metadata, but must never erase richer
        // articles or dynamic-field values already proven complete.
        var preserveArticles = existing?.ArticlesComplete == true && !profile.ArticlesComplete;
        var preserveDynamicFields = !profile.DynamicFieldsComplete && existing?.DynamicFieldsComplete == true;
        var merged = context with
        {
            Articles = preserveArticles ? existing!.Context.Articles : context.Articles,
            ReplySourceArticle = preserveArticles ? existing!.Context.ReplySourceArticle : context.ReplySourceArticle,
            ReplyRecipient = preserveArticles ? existing!.Context.ReplyRecipient : context.ReplyRecipient,
            CostCenterValue = preserveDynamicFields ? existing!.Context.CostCenterValue : context.CostCenterValue,
            OrderValue = preserveDynamicFields ? existing!.Context.OrderValue : context.OrderValue
        };
        var mergedProfile = new TicketDetailFetchProfile(
            profile.MetadataComplete || (sameRemoteVersion && existing?.MetadataComplete == true),
            profile.ArticlesComplete || (sameRemoteVersion && existing?.ArticlesComplete == true),
            profile.DynamicFieldsComplete || (sameRemoteVersion && existing?.DynamicFieldsComplete == true),
            profile.ArticlesComplete
                ? Math.Max(profile.FetchedArticleLimit, sameRemoteVersion ? existing?.FetchedArticleLimit ?? 0 : 0)
                : sameRemoteVersion ? existing?.FetchedArticleLimit ?? 0 : 0);

        using var connection = Open();
        using var transaction = connection.BeginTransaction();
        using (var command = connection.CreateCommand())
        {
            command.Transaction = transaction;
            command.CommandText = @"INSERT INTO znuny_ticket_detail_cache
(ticket_id,ticket_number,title,state,remote_changed_utc,last_fetched_utc,cost_center_value,order_value,reply_recipient,reply_source_article_id,metadata_complete,articles_complete,dynamic_fields_complete,fetched_article_limit)
VALUES($id,$number,$title,$state,$changed,$fetched,$cost,$order,$recipient,$reply,$metadata,$articles,$dynamic,$limit)
ON CONFLICT(ticket_id) DO UPDATE SET ticket_number=excluded.ticket_number,title=excluded.title,state=excluded.state,
remote_changed_utc=excluded.remote_changed_utc,last_fetched_utc=excluded.last_fetched_utc,cost_center_value=excluded.cost_center_value,
order_value=excluded.order_value,reply_recipient=excluded.reply_recipient,reply_source_article_id=excluded.reply_source_article_id,
metadata_complete=excluded.metadata_complete,articles_complete=excluded.articles_complete,
dynamic_fields_complete=excluded.dynamic_fields_complete,fetched_article_limit=excluded.fetched_article_limit";
            command.Parameters.AddWithValue("$id", merged.TicketId); command.Parameters.AddWithValue("$number", merged.TicketNumber);
            command.Parameters.AddWithValue("$title", merged.TicketTitle); command.Parameters.AddWithValue("$state", state);
            command.Parameters.AddWithValue("$changed", (object?)remoteChangedUtc?.ToUniversalTime().ToString("O") ?? DBNull.Value);
            command.Parameters.AddWithValue("$fetched", DateTime.UtcNow.ToString("O")); command.Parameters.AddWithValue("$cost", merged.CostCenterValue);
            command.Parameters.AddWithValue("$order", merged.OrderValue); command.Parameters.AddWithValue("$recipient", merged.ReplyRecipient);
            command.Parameters.AddWithValue("$reply", merged.ReplySourceArticle?.ArticleId ?? "");
            command.Parameters.AddWithValue("$metadata", mergedProfile.MetadataComplete ? 1 : 0);
            command.Parameters.AddWithValue("$articles", mergedProfile.ArticlesComplete ? 1 : 0);
            command.Parameters.AddWithValue("$dynamic", mergedProfile.DynamicFieldsComplete ? 1 : 0);
            command.Parameters.AddWithValue("$limit", mergedProfile.FetchedArticleLimit); command.ExecuteNonQuery();
        }
        if (!preserveArticles)
        {
            using var delete = connection.CreateCommand(); delete.Transaction = transaction;
            delete.CommandText = "DELETE FROM znuny_ticket_article_cache WHERE ticket_id=$id"; delete.Parameters.AddWithValue("$id", merged.TicketId); delete.ExecuteNonQuery();
            for (var index = 0; index < merged.Articles.Count; index++)
            {
                using var insert = connection.CreateCommand(); insert.Transaction = transaction;
                insert.CommandText = "INSERT INTO znuny_ticket_article_cache(ticket_id,article_id,ordinal,payload_json) VALUES($id,$article,$ordinal,$json)";
                insert.Parameters.AddWithValue("$id", merged.TicketId); insert.Parameters.AddWithValue("$article", merged.Articles[index].ArticleId);
                insert.Parameters.AddWithValue("$ordinal", index); insert.Parameters.AddWithValue("$json", JsonSerializer.Serialize(merged.Articles[index])); insert.ExecuteNonQuery();
            }
        }
        transaction.Commit();
    }

    public IReadOnlyList<TicketFieldOption> LoadFieldOptions(string fieldName, string fingerprint, TimeSpan ttl, bool allowExpired = false)
        => LoadFieldOptionsEntry(fieldName, fingerprint) is { } entry
           && (allowExpired || entry.IsFresh(ttl, DateTime.UtcNow)) ? entry.Options : Array.Empty<TicketFieldOption>();

    public DynamicFieldOptionsCacheEntry LoadFieldOptionsEntry(string fieldName, string fingerprint)
    {
        var result = new List<TicketFieldOption>(); using var connection = Open(); using var command = connection.CreateCommand();
        command.CommandText = @"SELECT option_key,display_value,fetched_utc FROM znuny_dynamic_field_options_cache
WHERE field_name=$field AND configuration_fingerprint=$fingerprint ORDER BY display_value";
        command.Parameters.AddWithValue("$field", fieldName); command.Parameters.AddWithValue("$fingerprint", fingerprint);
        using var reader = command.ExecuteReader();
        DateTime? oldestFetched = null;
        while (reader.Read())
        {
            var fetched = Parse(reader.GetString(2));
            if (fetched.HasValue && (!oldestFetched.HasValue || fetched < oldestFetched)) oldestFetched = fetched;
            result.Add(new TicketFieldOption(reader.GetString(0), reader.GetString(1)));
        }
        return new DynamicFieldOptionsCacheEntry(result, oldestFetched);
    }

    public void ReplaceFieldOptions(string fieldName, string fingerprint, IReadOnlyCollection<TicketFieldOption> options)
    {
        if (options.Count == 0) return; // a failed/empty response never becomes a valid cache
        using var connection = Open(); using var transaction = connection.BeginTransaction();
        using (var delete = connection.CreateCommand()) { delete.Transaction = transaction; delete.CommandText = "DELETE FROM znuny_dynamic_field_options_cache WHERE field_name=$field AND configuration_fingerprint=$fingerprint"; delete.Parameters.AddWithValue("$field", fieldName); delete.Parameters.AddWithValue("$fingerprint", fingerprint); delete.ExecuteNonQuery(); }
        foreach (var option in options)
        {
            using var insert = connection.CreateCommand(); insert.Transaction = transaction;
            insert.CommandText = @"INSERT INTO znuny_dynamic_field_options_cache(field_name,option_key,display_value,fetched_utc,configuration_fingerprint)
VALUES($field,$key,$display,$fetched,$fingerprint)";
            insert.Parameters.AddWithValue("$field", fieldName); insert.Parameters.AddWithValue("$key", option.Key); insert.Parameters.AddWithValue("$display", option.DisplayText);
            insert.Parameters.AddWithValue("$fetched", DateTime.UtcNow.ToString("O")); insert.Parameters.AddWithValue("$fingerprint", fingerprint); insert.ExecuteNonQuery();
        }
        transaction.Commit();
    }

    public string LoadSyncCursor(string contextKey)
    {
        using var connection = Open(); using var command = connection.CreateCommand();
        command.CommandText = "SELECT next_ticket_id FROM znuny_sync_cursor WHERE context_key=$key"; command.Parameters.AddWithValue("$key", contextKey);
        return command.ExecuteScalar()?.ToString() ?? string.Empty;
    }

    public void StoreSyncCursor(string contextKey, string nextTicketId)
    {
        using var connection = Open(); using var command = connection.CreateCommand();
        command.CommandText = @"INSERT INTO znuny_sync_cursor(context_key,next_ticket_id,updated_utc) VALUES($key,$id,$utc)
ON CONFLICT(context_key) DO UPDATE SET next_ticket_id=excluded.next_ticket_id,updated_utc=excluded.updated_utc";
        command.Parameters.AddWithValue("$key", contextKey); command.Parameters.AddWithValue("$id", nextTicketId); command.Parameters.AddWithValue("$utc", DateTime.UtcNow.ToString("O")); command.ExecuteNonQuery();
    }

    public TicketReconciliationCycle StartOrLoadCycle(string contextKey, string discoveryFingerprint,
        IReadOnlyList<string> discoveredTicketIds)
    {
        using var connection = Open(); using var transaction = connection.BeginTransaction();
        string? storedFingerprint = null; string? discoveredJson = null; DateTime? started = null;
        using (var load = connection.CreateCommand())
        {
            load.Transaction = transaction;
            load.CommandText = "SELECT discovery_fingerprint,discovered_ticket_ids_json,started_utc FROM znuny_reconciliation_cycle WHERE context_key=$key";
            load.Parameters.AddWithValue("$key", contextKey);
            using var reader = load.ExecuteReader();
            if (reader.Read()) { storedFingerprint = reader.GetString(0); discoveredJson = reader.GetString(1); started = Parse(reader.GetString(2)); }
        }
        if (!string.Equals(storedFingerprint, discoveryFingerprint, StringComparison.Ordinal))
        {
            var oldDiscovered = JsonSerializer.Deserialize<List<string>>(discoveredJson ?? "[]") ?? [];
            var oldPending = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var loadPending = connection.CreateCommand())
            {
                loadPending.Transaction = transaction;
                loadPending.CommandText = "SELECT ticket_id FROM znuny_reconciliation_pending WHERE context_key=$key";
                loadPending.Parameters.AddWithValue("$key", contextKey);
                using var reader = loadPending.ExecuteReader();
                while (reader.Read()) oldPending.Add(reader.GetString(0));
            }
            // IDs completed for the previous discovery and still present are safe to
            // retain. New IDs and removals (which callers include for verification)
            // are queued exactly once.
            var completedStillPresent = oldDiscovered.Except(oldPending, StringComparer.OrdinalIgnoreCase)
                .Intersect(discoveredTicketIds, StringComparer.OrdinalIgnoreCase).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var nextPending = discoveredTicketIds.Where(id => !completedStillPresent.Contains(id)).ToList();
            using (var clearPending = connection.CreateCommand())
            {
                clearPending.Transaction = transaction; clearPending.CommandText = "DELETE FROM znuny_reconciliation_pending WHERE context_key=$key";
                clearPending.Parameters.AddWithValue("$key", contextKey); clearPending.ExecuteNonQuery();
            }
            using var clear = connection.CreateCommand(); clear.Transaction = transaction;
            clear.CommandText = "DELETE FROM znuny_reconciliation_cycle WHERE context_key=$key"; clear.Parameters.AddWithValue("$key", contextKey); clear.ExecuteNonQuery();
            var now = DateTime.UtcNow;
            using var cycle = connection.CreateCommand(); cycle.Transaction = transaction;
            cycle.CommandText = "INSERT INTO znuny_reconciliation_cycle(context_key,discovery_fingerprint,discovered_ticket_ids_json,started_utc) VALUES($key,$fingerprint,$ids,$utc)";
            cycle.Parameters.AddWithValue("$key", contextKey); cycle.Parameters.AddWithValue("$fingerprint", discoveryFingerprint);
            cycle.Parameters.AddWithValue("$ids", JsonSerializer.Serialize(discoveredTicketIds)); cycle.Parameters.AddWithValue("$utc", now.ToString("O")); cycle.ExecuteNonQuery();
            for (var index = 0; index < nextPending.Count; index++)
            {
                using var pending = connection.CreateCommand(); pending.Transaction = transaction;
                pending.CommandText = "INSERT INTO znuny_reconciliation_pending(context_key,ticket_id,ordinal) VALUES($key,$id,$ordinal)";
                pending.Parameters.AddWithValue("$key", contextKey); pending.Parameters.AddWithValue("$id", nextPending[index]); pending.Parameters.AddWithValue("$ordinal", index); pending.ExecuteNonQuery();
            }
            storedFingerprint = discoveryFingerprint; discoveredJson = JsonSerializer.Serialize(discoveredTicketIds); started = now;
        }
        var pendingIds = new List<string>();
        using (var pending = connection.CreateCommand())
        {
            pending.Transaction = transaction; pending.CommandText = "SELECT ticket_id FROM znuny_reconciliation_pending WHERE context_key=$key ORDER BY ordinal";
            pending.Parameters.AddWithValue("$key", contextKey); using var reader = pending.ExecuteReader(); while (reader.Read()) pendingIds.Add(reader.GetString(0));
        }
        transaction.Commit();
        return new TicketReconciliationCycle(contextKey, storedFingerprint!,
            JsonSerializer.Deserialize<List<string>>(discoveredJson!) ?? [], pendingIds, started ?? DateTime.UtcNow);
    }

    public void CompleteCycleTicket(string contextKey, string ticketId)
    {
        using var connection = Open(); using var command = connection.CreateCommand();
        command.CommandText = "DELETE FROM znuny_reconciliation_pending WHERE context_key=$key AND ticket_id=$id";
        command.Parameters.AddWithValue("$key", contextKey); command.Parameters.AddWithValue("$id", ticketId); command.ExecuteNonQuery();
    }

    public void CompleteCycle(string contextKey)
    {
        using var connection = Open(); using var transaction = connection.BeginTransaction();
        using (var pending = connection.CreateCommand()) { pending.Transaction = transaction; pending.CommandText = "DELETE FROM znuny_reconciliation_pending WHERE context_key=$key"; pending.Parameters.AddWithValue("$key", contextKey); pending.ExecuteNonQuery(); }
        using var command = connection.CreateCommand(); command.Transaction = transaction;
        command.CommandText = "DELETE FROM znuny_reconciliation_cycle WHERE context_key=$key"; command.Parameters.AddWithValue("$key", contextKey); command.ExecuteNonQuery(); transaction.Commit();
    }

    private static IReadOnlyList<TicketArticleItem> LoadArticles(SqliteConnection connection, string ticketId)
    {
        var result = new List<TicketArticleItem>(); using var command = connection.CreateCommand();
        command.CommandText = "SELECT payload_json FROM znuny_ticket_article_cache WHERE ticket_id=$id ORDER BY ordinal"; command.Parameters.AddWithValue("$id", ticketId);
        using var reader = command.ExecuteReader(); while (reader.Read()) { var item = JsonSerializer.Deserialize<TicketArticleItem>(reader.GetString(0)); if (item != null) result.Add(item); }
        return result;
    }
    private SqliteConnection Open() { var connection = new SqliteConnection(_database.ConnectionString); connection.Open(); return connection; }
    private static string Text(SqliteDataReader reader, string name) => reader[name]?.ToString() ?? "";
    private static string? NullableText(SqliteDataReader reader, string name) => reader[name] is DBNull ? null : reader[name]?.ToString();
    private static bool Bool(SqliteDataReader reader, string name) => Convert.ToInt64(reader[name], CultureInfo.InvariantCulture) != 0;
    private static int Int(SqliteDataReader reader, string name) => Convert.ToInt32(reader[name], CultureInfo.InvariantCulture);
    private static DateTime? Parse(string? value) => DateTime.TryParse(value, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out var parsed) ? parsed.ToUniversalTime() : null;
    private static bool SameRemoteVersion(DateTime? existing, DateTime? incoming)
        => existing.HasValue && incoming.HasValue && existing.Value == incoming.Value;
}
