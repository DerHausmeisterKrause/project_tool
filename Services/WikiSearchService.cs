using System.Diagnostics;
using System.Text.Json;
using Microsoft.Data.Sqlite;
using TaskTool.Models;

namespace TaskTool.Services;

public sealed class WikiSearchService
{
    private readonly DatabaseService _database;
    private readonly SettingsService _settings;
    private readonly LoggerService _logger;
    private readonly WikiKeywordExtractor _extractor;
    private readonly Dictionary<string, IWikiProvider> _providers;
    private readonly SemaphoreSlim _parallelism = new(3);

    public WikiSearchService(DatabaseService database, SettingsService settings, LoggerService logger, IEnumerable<IWikiProvider>? providers = null, WikiKeywordExtractor? extractor = null, WikiVocabularyIndexService? vocabulary = null)
    {
        _database = database; _settings = settings; _logger = logger; _extractor = extractor ?? new WikiKeywordExtractor(vocabulary, logger);
        providers ??= new IWikiProvider[] { new ConfluenceDataCenterWikiProvider(settings), new ConfluenceCloudWikiProvider(settings), new GenericRestWikiProvider(settings), new XWikiProvider(settings) };
        _providers = providers.ToDictionary(x => x.ProviderType, StringComparer.OrdinalIgnoreCase);
    }

    public IReadOnlyList<WikiSearchResult> LoadResults(Guid taskId)
    {
        using var conn = new SqliteConnection(_database.ConnectionString); conn.Open(); using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT source_id,external_id,title,url,excerpt,relevance_score,matched_terms,provider_rank,last_modified_utc,searched_at_utc FROM task_wiki_results WHERE task_id=$task ORDER BY relevance_score DESC";
        cmd.Parameters.AddWithValue("$task", taskId.ToString()); using var reader = cmd.ExecuteReader(); var list = new List<WikiSearchResult>();
        while (reader.Read()) { var sourceId = reader.GetString(0); list.Add(new WikiSearchResult { SourceId = sourceId, SourceName = _settings.Current.WikiSources.FirstOrDefault(x => x.Id == sourceId)?.Name ?? sourceId, ExternalId = reader.GetString(1), Title = reader.GetString(2), Url = reader.GetString(3), Excerpt = reader.GetString(4), RelevanceScore = reader.GetDouble(5), MatchedTerms = reader.GetString(6), ProviderRank = reader.GetInt32(7), LastModifiedUtc = reader.IsDBNull(8) ? null : DateTime.Parse(reader.GetString(8)), SearchedAtUtc = DateTime.Parse(reader.GetString(9)) }); }
        return list;
    }

    public IReadOnlyList<string> LoadSearchTerms(Guid taskId)
    {
        using var connection = new SqliteConnection(_database.ConnectionString);
        connection.Open();
        using var command = connection.CreateCommand();
        command.CommandText = @"SELECT keywords_json
FROM task_wiki_search_runs
WHERE task_id=$task AND status='success'
ORDER BY searched_at_utc DESC, source_id";
        command.Parameters.AddWithValue("$task", taskId.ToString());
        using var reader = command.ExecuteReader();
        var serializedRuns = new List<string?>();
        while (reader.Read())
            serializedRuns.Add(reader.IsDBNull(0) ? null : reader.GetString(0));

        return WikiSearchTermPersistence.MergeSerialized(serializedRuns);
    }

    public async Task<WikiSearchSummary> SearchAsync(TaskItem task, string ticketTitle, string firstMessage, bool force, CancellationToken token = default)
    {
        if (!task.IsZnunyTask || !task.IsZnunyAssigned) return new(0, 0);
        var sources = new List<WikiSourceSettings>();
        foreach (var source in _settings.Current.WikiSources.Where(x => x.Enabled))
        {
            if (!WikiSourceValidation.TryValidate(source, out var configurationError))
            {
                _logger.Warning($"[WikiSearch] stage=configuration sourceId={Safe(source.Id)} provider={Safe(source.ProviderType)} host={Host(source.BaseUrl)} status=skipped message='{Sanitize(configurationError)}'");
                continue;
            }
            if (force || !HasRun(task.Id, source.Id)) sources.Add(source);
        }
        var outcomes = await Task.WhenAll(sources.Select(source => SearchSourceAsync(task.Id, source, ticketTitle, firstMessage, force, token)));
        return new(outcomes.Count(x => x), outcomes.Count(x => !x));
    }

    public async Task<IReadOnlyList<WikiProviderResult>> TestSourceAsync(WikiSourceSettings source, string searchTerm, CancellationToken token = default)
    {
        if (!WikiSourceValidation.TryValidate(source, out var error)) throw new InvalidOperationException(error);
        if (string.IsNullOrWhiteSpace(searchTerm)) throw new InvalidOperationException("Bitte einen Test-Suchbegriff eingeben.");
        if (!_providers.TryGetValue(source.ProviderType, out var provider)) throw new InvalidOperationException("Wiki-Quelle besitzt keinen gültigen Provider.");
        return await provider.SearchAsync(source, new[] { searchTerm.Trim() }, Math.Min(20, Math.Max(1, source.MaxResults)), token);
    }

    public void ResetFailedRunsForSource(string sourceId)
    {
        if (string.IsNullOrWhiteSpace(sourceId)) return;
        using var conn = new SqliteConnection(_database.ConnectionString); conn.Open(); using var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM task_wiki_search_runs WHERE source_id=$source AND status='failed'";
        cmd.Parameters.AddWithValue("$source", sourceId); cmd.ExecuteNonQuery();
    }

    public void InvalidateSource(string sourceId)
    {
        if (string.IsNullOrWhiteSpace(sourceId)) return;
        using var conn = new SqliteConnection(_database.ConnectionString); conn.Open(); using var tx = conn.BeginTransaction();
        foreach (var table in new[] { "task_wiki_results", "task_wiki_search_runs" }) { using var cmd = conn.CreateCommand(); cmd.Transaction = tx; cmd.CommandText = $"DELETE FROM {table} WHERE source_id=$source"; cmd.Parameters.AddWithValue("$source", sourceId); cmd.ExecuteNonQuery(); }
        tx.Commit();
    }

    private bool HasRun(Guid taskId, string sourceId)
    {
        using var conn = new SqliteConnection(_database.ConnectionString); conn.Open(); using var cmd = conn.CreateCommand(); cmd.CommandText = "SELECT 1 FROM task_wiki_search_runs WHERE task_id=$t AND source_id=$s LIMIT 1"; cmd.Parameters.AddWithValue("$t", taskId.ToString()); cmd.Parameters.AddWithValue("$s", sourceId); return cmd.ExecuteScalar() != null;
    }

    private async Task<bool> SearchSourceAsync(Guid taskId, WikiSourceSettings source, string title, string firstMessage, bool force, CancellationToken token)
    {
        await _parallelism.WaitAsync(token); var watch = Stopwatch.StartNew();
        var weightedTerms = _extractor.ExtractForSource(source, title, firstMessage); var terms = weightedTerms.Select(x => x.Text).ToList();
        try
        {
            if (terms.Count == 0) throw new InvalidOperationException("Keine geeigneten Suchbegriffe erkannt.");
            if (!WikiSourceValidation.TryValidate(source, out var configurationError)) throw new InvalidOperationException(configurationError);
            if (!_providers.TryGetValue(source.ProviderType, out var provider)) throw new InvalidOperationException("Wiki-Quelle besitzt keinen gültigen Provider.");
            var raw = await provider.SearchAsync(source, terms, Math.Min(40, Math.Max(source.MaxResults * 4, 20)), token);
            var ranked = raw.GroupBy(x => x.Url, StringComparer.OrdinalIgnoreCase).Select(x => x.First()).Select(x => Rank(x, weightedTerms, title)).OrderByDescending(x => x.RelevanceScore).ThenBy(x => x.ProviderRank).Take(source.MaxResults).ToList();
            StoreSuccess(taskId, source.Id, terms, ranked);
            _logger.Info($"[WikiSearch] taskId={taskId} sourceId={source.Id} provider={source.ProviderType} termCount={terms.Count} resultCount={raw.Count} storedCount={ranked.Count} durationMs={watch.ElapsedMilliseconds} status=success"); return true;
        }
        catch (OperationCanceledException) when (token.IsCancellationRequested) { throw; }
        catch (Exception ex)
        {
            StoreFailure(taskId, source.Id, terms, ex.GetType().Name, force);
            _logger.Warning($"[WikiSearch] taskId={taskId} stage=request sourceId={Safe(source.Id)} provider={Safe(source.ProviderType)} host={Host(source.BaseUrl)} termCount={terms.Count} durationMs={watch.ElapsedMilliseconds} status=failed errorType={ex.GetType().Name} message='{Sanitize(ex.Message)}'"); return false;
        }
        finally { _parallelism.Release(); }
    }

    private static WikiSearchResult Rank(WikiProviderResult item, IReadOnlyList<WikiSearchTerm> terms, string ticketTitle)
    {
        var matched = new List<string>(); double score = Math.Max(0, 8 - item.ProviderRank) * .5;
        foreach (var term in terms) { var inTitle = item.Title.Contains(term.Text, StringComparison.OrdinalIgnoreCase); var inExcerpt = item.Excerpt.Contains(term.Text, StringComparison.OrdinalIgnoreCase); if (inTitle || inExcerpt) matched.Add(term.Text); if (inTitle) score += 20 + term.Score * .35 + (ticketTitle.Contains(term.Text, StringComparison.OrdinalIgnoreCase) ? 15 : 0); else if (inExcerpt) score += 3 + term.Score * .08; }
        if (matched.Count > 1) score += (matched.Count - 1) * 8;
        return new WikiSearchResult { ExternalId = item.ExternalId, Title = item.Title, Url = item.Url, Excerpt = item.Excerpt, ProviderRank = item.ProviderRank, LastModifiedUtc = item.LastModifiedUtc, MatchedTerms = string.Join(" · ", matched), RelevanceScore = Math.Clamp(score, 0, 100), SearchedAtUtc = DateTime.UtcNow };
    }

    private void StoreSuccess(Guid taskId, string sourceId, IReadOnlyList<string> terms, IReadOnlyList<WikiSearchResult> results)
    {
        using var conn = new SqliteConnection(_database.ConnectionString); conn.Open(); using var tx = conn.BeginTransaction();
        Execute(conn, tx, "DELETE FROM task_wiki_results WHERE task_id=$t AND source_id=$s", taskId, sourceId);
        foreach (var r in results) { using var cmd = conn.CreateCommand(); cmd.Transaction = tx; cmd.CommandText = "INSERT INTO task_wiki_results(task_id,source_id,external_id,title,url,excerpt,relevance_score,matched_terms,provider_rank,last_modified_utc,searched_at_utc) VALUES($t,$s,$i,$n,$u,$e,$r,$m,$p,$l,$a)"; cmd.Parameters.AddWithValue("$t", taskId.ToString()); cmd.Parameters.AddWithValue("$s", sourceId); cmd.Parameters.AddWithValue("$i", r.ExternalId); cmd.Parameters.AddWithValue("$n", r.Title); cmd.Parameters.AddWithValue("$u", r.Url); cmd.Parameters.AddWithValue("$e", r.Excerpt); cmd.Parameters.AddWithValue("$r", r.RelevanceScore); cmd.Parameters.AddWithValue("$m", r.MatchedTerms); cmd.Parameters.AddWithValue("$p", r.ProviderRank); cmd.Parameters.AddWithValue("$l", r.LastModifiedUtc?.ToString("O") ?? (object)DBNull.Value); cmd.Parameters.AddWithValue("$a", r.SearchedAtUtc.ToString("O")); cmd.ExecuteNonQuery(); }
        UpsertRun(conn, tx, taskId, sourceId, terms, "success", ""); tx.Commit();
    }
    private void StoreFailure(Guid taskId, string sourceId, IReadOnlyList<string> terms, string error, bool force)
    { using var conn = new SqliteConnection(_database.ConnectionString); conn.Open(); using var tx = conn.BeginTransaction(); UpsertRun(conn, tx, taskId, sourceId, terms, "failed", error); tx.Commit(); }
    private static void UpsertRun(SqliteConnection conn, SqliteTransaction tx, Guid taskId, string sourceId, IReadOnlyList<string> terms, string status, string error) { using var cmd = conn.CreateCommand(); cmd.Transaction = tx; cmd.CommandText = "INSERT INTO task_wiki_search_runs(task_id,source_id,searched_at_utc,status,keywords_json,error) VALUES($t,$s,$a,$x,$k,$e) ON CONFLICT(task_id,source_id) DO UPDATE SET searched_at_utc=excluded.searched_at_utc,status=excluded.status,keywords_json=excluded.keywords_json,error=excluded.error"; cmd.Parameters.AddWithValue("$t", taskId.ToString()); cmd.Parameters.AddWithValue("$s", sourceId); cmd.Parameters.AddWithValue("$a", DateTime.UtcNow.ToString("O")); cmd.Parameters.AddWithValue("$x", status); cmd.Parameters.AddWithValue("$k", JsonSerializer.Serialize(terms)); cmd.Parameters.AddWithValue("$e", error); cmd.ExecuteNonQuery(); }
    private static void Execute(SqliteConnection conn, SqliteTransaction tx, string sql, Guid taskId, string sourceId) { using var cmd = conn.CreateCommand(); cmd.Transaction = tx; cmd.CommandText = sql; cmd.Parameters.AddWithValue("$t", taskId.ToString()); cmd.Parameters.AddWithValue("$s", sourceId); cmd.ExecuteNonQuery(); }
    private static string Safe(string? value) => string.IsNullOrWhiteSpace(value) ? "missing" : value.Replace(" ", "_");
    private static string Host(string? value) => Uri.TryCreate(value, UriKind.Absolute, out var uri) ? uri.Host : "invalid";
    private static string Sanitize(string? value)
    {
        var sanitized = (value ?? string.Empty).Replace("\r", " ").Replace("\n", " ").Replace("'", "");
        return sanitized.Length > 240 ? sanitized[..240] : sanitized;
    }
}
