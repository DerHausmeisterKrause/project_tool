using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Data.Sqlite;
using TaskTool.Models;

namespace TaskTool.Services;

public sealed record AiRetrievalMatch(AiKnowledgeSourceType SourceType, string Content, string Title, string DisplaySource, string? Url, double Score, int? PageNumber = null);
public enum AiKnowledgeSourceType { LocalFiles, Wiki }

public sealed class WikiAiKnowledgeService : IDisposable
{
    public static readonly TimeSpan SyncInterval = TimeSpan.FromMinutes(30);
    private readonly SettingsService _settings;
    private readonly LoggerService _logger;
    private readonly Dictionary<string, IWikiKnowledgeProvider> _providers;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private readonly Timer _timer;
    public string IndexPath { get; }
    public event EventHandler? StatusChanged;

    public WikiAiKnowledgeService(SettingsService settings, LoggerService logger, IEnumerable<(string Type, IWikiKnowledgeProvider Provider)>? providers = null, string? localAppData = null)
    {
        _settings = settings; _logger = logger;
        _providers = providers?.ToDictionary(x => x.Type, x => x.Provider, StringComparer.OrdinalIgnoreCase)
            ?? new IWikiKnowledgeProvider[] { new ConfluenceDataCenterWikiProvider(settings), new ConfluenceCloudWikiProvider(settings) }
                .ToDictionary(x => ((IWikiProvider)x).ProviderType, StringComparer.OrdinalIgnoreCase);
        var root = localAppData ?? Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var directory = Path.Combine(root, "Plenaro", "AI"); Directory.CreateDirectory(directory);
        IndexPath = Path.Combine(directory, "wiki-knowledge-index.db");
        Initialize();
        _timer = new Timer(_ => _ = SyncStaleAsync(), null, SyncInterval, SyncInterval);
    }

    public bool IsAvailable => _settings.Current.WikiSources.Any(IsIndexable);
    public bool IsIndexable(WikiSourceSettings source) => source.Enabled && WikiScopePolicy.SupportsAiKnowledge(source) && _providers.ContainsKey(source.ProviderType) && WikiSourceValidation.TryValidate(source, out _);

    public async Task SyncStaleAsync(CancellationToken token = default)
    {
        foreach (var source in _settings.Current.WikiSources.Where(IsIndexable))
        {
            var status = GetStatus(source.Id);
            if (status.LastSuccessUtc is null || DateTime.UtcNow - status.LastSuccessUtc >= SyncInterval)
                await SyncAsync(source, false, token);
        }
    }

    public async Task SyncAsync(WikiSourceSettings source, bool fullRebuild, CancellationToken token = default)
    {
        if (!IsIndexable(source) || !await _gate.WaitAsync(0, token)) return;
        var watch = Stopwatch.StartNew(); var fingerprint = WikiScopePolicy.Fingerprint(source);
        try
        {
            var old = LoadPageMetadata(source.Id);
            var previousFingerprint = GetFingerprint(source.Id);
            fullRebuild |= old.Count == 0 || !string.Equals(previousFingerprint, fingerprint, StringComparison.Ordinal);
            SetTransientStatus(source.Id, fullRebuild ? "full-index" : "syncing");
            _logger.Info($"[Wiki AI Index] sourceId={source.Id} action=sync-start mode={(fullRebuild ? "full" : "incremental")}");
            var pages = new List<WikiKnowledgePage>(); const int size = 100; var offset = 0;
            while (true)
            {
                var batch = await _providers[source.ProviderType].GetPagesAsync(source, offset, size, token);
                pages.AddRange(batch.Pages); offset += batch.Pages.Count;
                if (!batch.HasMore || batch.Pages.Count == 0) break;
            }
            var seen = pages.ToDictionary(x => x.ExternalId, StringComparer.Ordinal);
            var changed = pages.Where(p => fullRebuild || !old.TryGetValue(p.ExternalId, out var existing) || IsChanged(existing, p)).ToArray();
            var loaded = new Dictionary<string, WikiKnowledgePageContent>(StringComparer.Ordinal);
            using var throttle = new SemaphoreSlim(3, 3);
            await Task.WhenAll(changed.Select(async page => { await throttle.WaitAsync(token); try { lock (loaded) loaded[page.ExternalId] = await _providers[source.ProviderType].GetPageContentAsync(source, page.ExternalId, token); } finally { throttle.Release(); } }));
            Apply(source, fingerprint, pages, loaded, fullRebuild);
            var deleted = old.Keys.Count(x => !seen.ContainsKey(x)); var added = changed.Count(x => !old.ContainsKey(x.ExternalId));
            _logger.Info($"[Wiki AI Index] sourceId={source.Id} pagesSeen={pages.Count} changed={changed.Length - added} new={added} deleted={deleted} unchanged={pages.Count - changed.Length}");
            var status = GetStatus(source.Id);
            _logger.Info($"[Wiki AI Index] sourceId={source.Id} action={(fullRebuild ? "full-index" : "sync-complete")} pages={status.PageCount} chunks={status.ChunkCount} durationMs={watch.ElapsedMilliseconds}");
        }
        catch (Exception ex) when (ex is not OperationCanceledException)
        {
            MarkFailed(source.Id); _logger.Warning($"[Wiki AI Index] sourceId={source.Id} action=sync-failed durationMs={watch.ElapsedMilliseconds} errorType={ex.GetType().Name}");
        }
        finally { _gate.Release(); StatusChanged?.Invoke(this, EventArgs.Empty); }
    }

    public async Task<IReadOnlyList<AiRetrievalMatch>> SearchAsync(string question, int limit = 10, CancellationToken token = default)
    {
        var terms = AiKnowledgeSearchService.NormalizeTerms(question); if (terms.Length == 0 || !File.Exists(IndexPath)) return Array.Empty<AiRetrievalMatch>();
        var watch = Stopwatch.StartNew(); var result = new List<AiRetrievalMatch>(); var candidates = 0;
        await using var db = Open(true); await using var cmd = db.CreateCommand();
        cmd.CommandText = "SELECT c.content,c.title,s.name,c.url,c.space_key,bm25(wiki_ai_chunks_fts,1,4,3) FROM wiki_ai_chunks_fts f JOIN wiki_ai_chunks c ON c.id=f.rowid JOIN wiki_ai_source_names s ON s.source_id=c.source_id WHERE wiki_ai_chunks_fts MATCH $q LIMIT $l";
        cmd.Parameters.AddWithValue("$q", string.Join(" OR ", terms.Select(x => $"\"{x}\"*"))); cmd.Parameters.AddWithValue("$l", limit * 4);
        await using var reader = await cmd.ExecuteReaderAsync(token);
        while (await reader.ReadAsync(token)) { candidates++; var content = reader.GetString(0); var title = reader.GetString(1); var score = Score(terms, content, title, reader.GetString(4)); if (score > 0) result.Add(new(AiKnowledgeSourceType.Wiki, content, title, reader.GetString(2), reader.GetString(3), score)); }
        var accepted = result.OrderByDescending(x => x.Score).Take(limit).ToArray();
        _logger.OperationalInfo($"[AI Wiki] Query completed candidates={candidates} accepted={accepted.Length} durationMs={watch.ElapsedMilliseconds}"); return accepted;
    }

    public WikiAiIndexStatus GetStatus(string sourceId)
    {
        using var db = Open(); using var cmd = db.CreateCommand(); cmd.CommandText = "SELECT status,page_count,chunk_count,last_success_utc FROM wiki_ai_sources WHERE source_id=$s"; cmd.Parameters.AddWithValue("$s", sourceId);
        using var r = cmd.ExecuteReader(); return r.Read() ? new(sourceId, r.GetInt32(1), r.GetInt32(2), r.IsDBNull(3) ? null : DateTime.Parse(r.GetString(3)).ToUniversalTime(), r.GetString(0)) : new(sourceId, 0, 0, null, "not-indexed");
    }

    public void Invalidate(string sourceId) { using var db = Open(); using var cmd = db.CreateCommand(); cmd.CommandText = "UPDATE wiki_ai_sources SET scope_fingerprint='' WHERE source_id=$s"; cmd.Parameters.AddWithValue("$s", sourceId); cmd.ExecuteNonQuery(); }
    private Dictionary<string, (string Version, DateTime? Modified)> LoadPageMetadata(string id) { using var db = Open(); using var cmd = db.CreateCommand(); cmd.CommandText="SELECT external_id,version,last_modified_utc FROM wiki_ai_pages WHERE source_id=$s"; cmd.Parameters.AddWithValue("$s",id); using var r=cmd.ExecuteReader(); var d=new Dictionary<string,(string,DateTime?)>(); while(r.Read()) d[r.GetString(0)]=(r.GetString(1),r.IsDBNull(2)?null:DateTime.Parse(r.GetString(2)).ToUniversalTime()); return d; }
    private string? GetFingerprint(string id) { using var db=Open(); using var cmd=db.CreateCommand(); cmd.CommandText="SELECT scope_fingerprint FROM wiki_ai_sources WHERE source_id=$s";cmd.Parameters.AddWithValue("$s",id);return cmd.ExecuteScalar() as string; }
    private static bool IsChanged((string Version, DateTime? Modified) old, WikiKnowledgePage page) => (!string.IsNullOrEmpty(page.Version) && old.Version != page.Version) || (page.LastModifiedUtc.HasValue && old.Modified != page.LastModifiedUtc);
    private void Apply(WikiSourceSettings source,string fingerprint,IReadOnlyList<WikiKnowledgePage> pages,IReadOnlyDictionary<string,WikiKnowledgePageContent> loaded,bool full)
    {
        using var db=Open(); using var tx=db.BeginTransaction();
        var ids=pages.Select(x=>x.ExternalId).ToHashSet(StringComparer.Ordinal); using(var q=db.CreateCommand()){q.Transaction=tx;q.CommandText="SELECT external_id FROM wiki_ai_pages WHERE source_id=$s";q.Parameters.AddWithValue("$s",source.Id);using var r=q.ExecuteReader();var deleted=new List<string>();while(r.Read())if(!ids.Contains(r.GetString(0)))deleted.Add(r.GetString(0));r.Close();foreach(var id in deleted)DeletePage(db,tx,source.Id,id);}
        foreach(var p in pages.Where(x=>loaded.ContainsKey(x.ExternalId))) { DeletePage(db,tx,source.Id,p.ExternalId); var body=loaded[p.ExternalId]; using var cmd=db.CreateCommand();cmd.Transaction=tx;cmd.CommandText="INSERT INTO wiki_ai_pages(source_id,external_id,title,url,space_key,version,last_modified_utc,content_hash,indexed_utc) VALUES($s,$e,$t,$u,$k,$v,$m,$h,$i)";cmd.Parameters.AddWithValue("$s",source.Id);cmd.Parameters.AddWithValue("$e",p.ExternalId);cmd.Parameters.AddWithValue("$t",p.Title);cmd.Parameters.AddWithValue("$u",p.Url);cmd.Parameters.AddWithValue("$k",p.SpaceKey);cmd.Parameters.AddWithValue("$v",body.Version);cmd.Parameters.AddWithValue("$m",body.LastModifiedUtc?.ToString("O")??(object)DBNull.Value);cmd.Parameters.AddWithValue("$h",Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(body.PlainText))));cmd.Parameters.AddWithValue("$i",DateTime.UtcNow.ToString("O"));cmd.ExecuteNonQuery(); foreach(var (chunk,index) in Chunk(body.PlainText).Select((c,i)=>(c,i))){using var c=db.CreateCommand();c.Transaction=tx;c.CommandText="INSERT INTO wiki_ai_chunks(page_id,chunk_index,content,title,space_key,source_id,url,external_id) VALUES((SELECT id FROM wiki_ai_pages WHERE source_id=$s AND external_id=$e),$n,$c,$t,$k,$s,$u,$e)";c.Parameters.AddWithValue("$s",source.Id);c.Parameters.AddWithValue("$e",p.ExternalId);c.Parameters.AddWithValue("$n",index);c.Parameters.AddWithValue("$c",chunk);c.Parameters.AddWithValue("$t",p.Title);c.Parameters.AddWithValue("$k",p.SpaceKey);c.Parameters.AddWithValue("$u",p.Url);c.ExecuteNonQuery();}}
        using(var n=db.CreateCommand()){n.Transaction=tx;n.CommandText="INSERT OR REPLACE INTO wiki_ai_source_names(source_id,name) VALUES($s,$n)";n.Parameters.AddWithValue("$s",source.Id);n.Parameters.AddWithValue("$n",source.Name);n.ExecuteNonQuery();}
        using(var state=db.CreateCommand()){state.Transaction=tx;state.CommandText="INSERT INTO wiki_ai_sources(source_id,scope_fingerprint,last_full_sync_utc,last_incremental_sync_utc,last_success_utc,status,page_count,chunk_count) VALUES($s,$f,$full,$inc,$now,'current',(SELECT count(*) FROM wiki_ai_pages WHERE source_id=$s),(SELECT count(*) FROM wiki_ai_chunks WHERE source_id=$s)) ON CONFLICT(source_id) DO UPDATE SET scope_fingerprint=$f,last_full_sync_utc=COALESCE($full,last_full_sync_utc),last_incremental_sync_utc=$inc,last_success_utc=$now,status='current',page_count=(SELECT count(*) FROM wiki_ai_pages WHERE source_id=$s),chunk_count=(SELECT count(*) FROM wiki_ai_chunks WHERE source_id=$s)";state.Parameters.AddWithValue("$s",source.Id);state.Parameters.AddWithValue("$f",fingerprint);state.Parameters.AddWithValue("$full",full?DateTime.UtcNow.ToString("O"):(object)DBNull.Value);state.Parameters.AddWithValue("$inc",DateTime.UtcNow.ToString("O"));state.Parameters.AddWithValue("$now",DateTime.UtcNow.ToString("O"));state.ExecuteNonQuery();} tx.Commit();
    }
    private static void DeletePage(SqliteConnection db,SqliteTransaction tx,string source,string external){using var c=db.CreateCommand();c.Transaction=tx;c.CommandText="DELETE FROM wiki_ai_chunks WHERE source_id=$s AND external_id=$e; DELETE FROM wiki_ai_pages WHERE source_id=$s AND external_id=$e";c.Parameters.AddWithValue("$s",source);c.Parameters.AddWithValue("$e",external);c.ExecuteNonQuery();}
    private static IEnumerable<string> Chunk(string text){const int max=1200,overlap=180;for(var start=0;start<text.Length;){var length=Math.Min(max,text.Length-start);if(start+length<text.Length){var boundary=text.LastIndexOf('\n',start+length-1,length);if(boundary>start+600)length=boundary-start;}var value=text.Substring(start,length).Trim();if(value.Length>0)yield return value;if(start+length>=text.Length)break;start+=Math.Max(1,length-overlap);}}
    private static double Score(IReadOnlyList<string> terms,string content,string title,string space){var matches=terms.Count(t=>content.Contains(t,StringComparison.OrdinalIgnoreCase)||title.Contains(t,StringComparison.OrdinalIgnoreCase)||space.Contains(t,StringComparison.OrdinalIgnoreCase));if(matches<Math.Min(2,terms.Count))return 0;return matches*5+terms.Count(t=>title.Contains(t,StringComparison.OrdinalIgnoreCase))*4;}
    private void SetTransientStatus(string id,string status){using var db=Open();using var c=db.CreateCommand();c.CommandText="INSERT INTO wiki_ai_sources(source_id,scope_fingerprint,status,page_count,chunk_count) VALUES($s,'',$x,0,0) ON CONFLICT(source_id) DO UPDATE SET status=$x";c.Parameters.AddWithValue("$s",id);c.Parameters.AddWithValue("$x",status);c.ExecuteNonQuery();StatusChanged?.Invoke(this,EventArgs.Empty);}
    private void MarkFailed(string id){using var db=Open();using var c=db.CreateCommand();c.CommandText="INSERT INTO wiki_ai_sources(source_id,scope_fingerprint,status,page_count,chunk_count) VALUES($s,'','failed',0,0) ON CONFLICT(source_id) DO UPDATE SET status='failed'";c.Parameters.AddWithValue("$s",id);c.ExecuteNonQuery();}
    private SqliteConnection Open(bool readOnly=false){var db=new SqliteConnection($"Data Source={IndexPath}{(readOnly?";Mode=ReadOnly":"")}");db.Open();return db;}
    private void Initialize(){using var db=Open();using var cmd=db.CreateCommand();cmd.CommandText="""
CREATE TABLE IF NOT EXISTS wiki_ai_sources(source_id TEXT PRIMARY KEY,scope_fingerprint TEXT NOT NULL,last_full_sync_utc TEXT,last_incremental_sync_utc TEXT,last_success_utc TEXT,status TEXT NOT NULL,page_count INTEGER NOT NULL,chunk_count INTEGER NOT NULL);
CREATE TABLE IF NOT EXISTS wiki_ai_source_names(source_id TEXT PRIMARY KEY,name TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS wiki_ai_pages(id INTEGER PRIMARY KEY,source_id TEXT NOT NULL,external_id TEXT NOT NULL,title TEXT NOT NULL,url TEXT NOT NULL,space_key TEXT NOT NULL,version TEXT NOT NULL,last_modified_utc TEXT,content_hash TEXT NOT NULL,indexed_utc TEXT NOT NULL,UNIQUE(source_id,external_id));
CREATE TABLE IF NOT EXISTS wiki_ai_chunks(id INTEGER PRIMARY KEY,page_id INTEGER NOT NULL,chunk_index INTEGER NOT NULL,content TEXT NOT NULL,title TEXT NOT NULL,space_key TEXT NOT NULL,source_id TEXT NOT NULL,url TEXT NOT NULL,external_id TEXT NOT NULL,FOREIGN KEY(page_id) REFERENCES wiki_ai_pages(id) ON DELETE CASCADE);
CREATE VIRTUAL TABLE IF NOT EXISTS wiki_ai_chunks_fts USING fts5(content,title,space_key,content='wiki_ai_chunks',content_rowid='id');
CREATE TRIGGER IF NOT EXISTS wiki_ai_chunks_ai AFTER INSERT ON wiki_ai_chunks BEGIN INSERT INTO wiki_ai_chunks_fts(rowid,content,title,space_key) VALUES(new.id,new.content,new.title,new.space_key); END;
CREATE TRIGGER IF NOT EXISTS wiki_ai_chunks_ad AFTER DELETE ON wiki_ai_chunks BEGIN INSERT INTO wiki_ai_chunks_fts(wiki_ai_chunks_fts,rowid,content,title,space_key) VALUES('delete',old.id,old.content,old.title,old.space_key); END;
""";cmd.ExecuteNonQuery();}
    public void Dispose(){_timer.Dispose();_gate.Dispose();}
}
