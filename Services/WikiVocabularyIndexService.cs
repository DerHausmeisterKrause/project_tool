using System.Diagnostics;
using Microsoft.Data.Sqlite;
using TaskTool.Models;

namespace TaskTool.Services;

public sealed class WikiVocabularyIndexService
{
    private readonly DatabaseService _database; private readonly SettingsService _settings; private readonly LoggerService _logger;
    private readonly Dictionary<string, IWikiVocabularyProvider> _providers; private readonly SemaphoreSlim _refreshGate = new(1, 1);
    public WikiVocabularyIndexService(DatabaseService database, SettingsService settings, LoggerService logger)
    {
        _database = database; _settings = settings; _logger = logger;
        _providers = new IWikiVocabularyProvider[] { new ConfluenceDataCenterWikiProvider(settings), new ConfluenceCloudWikiProvider(settings) }.ToDictionary(x => ((IWikiProvider)x).ProviderType, StringComparer.OrdinalIgnoreCase);
    }

    public async Task RefreshStaleAsync(CancellationToken token = default)
    {
        foreach (var source in _settings.Current.WikiSources.Where(x => x.Enabled && WikiScopePolicy.SupportsSecureVocabulary(x)))
        { var status = GetStatus(source); if (status.Status != "success" || status.UpdatedUtc is not DateTime updated || DateTime.UtcNow - updated > TimeSpan.FromHours(24)) await RefreshAsync(source, token); }
    }

    public async Task RefreshAsync(WikiSourceSettings source, CancellationToken token = default)
    {
        if (!WikiScopePolicy.SupportsSecureVocabulary(source) || !_providers.TryGetValue(source.ProviderType, out var provider)) return;
        var fingerprint = WikiScopePolicy.Fingerprint(source); var watch = Stopwatch.StartNew(); await _refreshGate.WaitAsync(token);
        try
        {
            var pages = new List<WikiVocabularyPage>(); const int pageSize = 100; var offset = 0;
            while (true) { var batch = await provider.GetVocabularyPageAsync(source, offset, pageSize, token); pages.AddRange(batch.Pages); if (!batch.HasMore) break; offset += batch.Pages.Count; if (batch.Pages.Count == 0) break; }
            Store(source.Id, fingerprint, pages); _logger.Info($"[WikiVocabulary] sourceId={source.Id} pageCount={pages.Count} durationMs={watch.ElapsedMilliseconds} status=success");
        }
        catch (Exception ex) when (ex is not OperationCanceledException) { StoreFailure(source.Id, fingerprint, ex.GetType().Name); _logger.Warning($"[WikiVocabulary] sourceId={source.Id} pageCount=0 durationMs={watch.ElapsedMilliseconds} status=failed errorType={ex.GetType().Name}"); }
        finally { _refreshGate.Release(); }
    }

    public IReadOnlyList<WikiVocabularyPage> FindCandidates(WikiSourceSettings source, IEnumerable<string> seedTerms, int limit = 60)
    {
        var fingerprint = WikiScopePolicy.Fingerprint(source); var tokens = seedTerms.SelectMany(WikiKeywordExtractor.Tokenize).Select(WikiKeywordExtractor.Normalize).Where(x => x.Length > 2).Distinct().Take(12).ToArray(); if (tokens.Length == 0) return Array.Empty<WikiVocabularyPage>();
        using var conn = new SqliteConnection(_database.ConnectionString); conn.Open(); using var cmd = conn.CreateCommand();
        cmd.CommandText = "SELECT external_id,title,'' FROM wiki_vocabulary_fts WHERE source_id=$s AND scope_fingerprint=$f AND wiki_vocabulary_fts MATCH $q ORDER BY bm25(wiki_vocabulary_fts) LIMIT $l";
        cmd.Parameters.AddWithValue("$s", source.Id); cmd.Parameters.AddWithValue("$f", fingerprint); cmd.Parameters.AddWithValue("$q", string.Join(" OR ", tokens.Select(x => '"' + x.Replace("\"", "\"\"") + '"'))); cmd.Parameters.AddWithValue("$l", limit);
        var result = new List<WikiVocabularyPage>(); using var reader = cmd.ExecuteReader(); while (reader.Read()) result.Add(new(reader.GetString(0), reader.GetString(1), reader.GetString(2), "")); return result;
    }

    public WikiVocabularyStatus GetStatus(WikiSourceSettings source)
    {
        using var conn = new SqliteConnection(_database.ConnectionString); conn.Open(); using var cmd = conn.CreateCommand(); cmd.CommandText = "SELECT page_count,updated_utc,status FROM wiki_vocabulary_state WHERE source_id=$s AND scope_fingerprint=$f"; cmd.Parameters.AddWithValue("$s", source.Id); cmd.Parameters.AddWithValue("$f", WikiScopePolicy.Fingerprint(source)); using var r = cmd.ExecuteReader(); return r.Read() ? new(r.GetInt32(0), r.IsDBNull(1) ? null : DateTime.Parse(r.GetString(1)), r.GetString(2)) : new(0, null, "missing");
    }

    public void Invalidate(string sourceId) { using var conn = new SqliteConnection(_database.ConnectionString); conn.Open(); using var tx = conn.BeginTransaction(); foreach (var table in new[] { "wiki_vocabulary_pages", "wiki_vocabulary_fts", "wiki_vocabulary_state" }) { using var cmd = conn.CreateCommand(); cmd.Transaction = tx; cmd.CommandText = $"DELETE FROM {table} WHERE source_id=$s"; cmd.Parameters.AddWithValue("$s", sourceId); cmd.ExecuteNonQuery(); } tx.Commit(); }
    private void Store(string sourceId, string fingerprint, IReadOnlyList<WikiVocabularyPage> pages) { using var conn = new SqliteConnection(_database.ConnectionString); conn.Open(); using var tx = conn.BeginTransaction(); foreach (var table in new[] { "wiki_vocabulary_pages", "wiki_vocabulary_fts" }) { using var d = conn.CreateCommand(); d.Transaction = tx; d.CommandText = $"DELETE FROM {table} WHERE source_id=$s AND scope_fingerprint=$f"; d.Parameters.AddWithValue("$s", sourceId); d.Parameters.AddWithValue("$f", fingerprint); d.ExecuteNonQuery(); } foreach (var p in pages) { var normalized = WikiKeywordExtractor.Normalize(p.Title); using var a = conn.CreateCommand(); a.Transaction = tx; a.CommandText = "INSERT INTO wiki_vocabulary_pages VALUES($s,$f,$i,$k,$t,$n,$u,$m,$a)"; a.Parameters.AddWithValue("$s", sourceId); a.Parameters.AddWithValue("$f", fingerprint); a.Parameters.AddWithValue("$i", p.ExternalId); a.Parameters.AddWithValue("$k", p.SpaceKey); a.Parameters.AddWithValue("$t", p.Title); a.Parameters.AddWithValue("$n", normalized); a.Parameters.AddWithValue("$u", p.Url); a.Parameters.AddWithValue("$m", p.LastModifiedUtc?.ToString("O") ?? (object)DBNull.Value); a.Parameters.AddWithValue("$a", DateTime.UtcNow.ToString("O")); a.ExecuteNonQuery(); using var f = conn.CreateCommand(); f.Transaction = tx; f.CommandText = "INSERT INTO wiki_vocabulary_fts VALUES($s,$p,$i,$t,$n)"; f.Parameters.AddWithValue("$s", sourceId); f.Parameters.AddWithValue("$p", fingerprint); f.Parameters.AddWithValue("$i", p.ExternalId); f.Parameters.AddWithValue("$t", p.Title); f.Parameters.AddWithValue("$n", normalized); f.ExecuteNonQuery(); } UpsertState(conn, tx, sourceId, fingerprint, pages.Count, "success", ""); tx.Commit(); }
    private void StoreFailure(string sourceId, string fingerprint, string error) { using var c = new SqliteConnection(_database.ConnectionString); c.Open(); using var t = c.BeginTransaction(); UpsertState(c, t, sourceId, fingerprint, 0, "failed", error); t.Commit(); }
    private static void UpsertState(SqliteConnection c, SqliteTransaction t, string s, string f, int count, string status, string error) { using var cmd = c.CreateCommand(); cmd.Transaction = t; cmd.CommandText = "INSERT INTO wiki_vocabulary_state VALUES($s,$f,$c,$u,$x,$e) ON CONFLICT(source_id,scope_fingerprint) DO UPDATE SET page_count=$c,updated_utc=$u,status=$x,error=$e"; cmd.Parameters.AddWithValue("$s", s); cmd.Parameters.AddWithValue("$f", f); cmd.Parameters.AddWithValue("$c", count); cmd.Parameters.AddWithValue("$u", DateTime.UtcNow.ToString("O")); cmd.Parameters.AddWithValue("$x", status); cmd.Parameters.AddWithValue("$e", error); cmd.ExecuteNonQuery(); }
}
