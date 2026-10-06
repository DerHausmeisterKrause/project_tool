using System.Diagnostics;
using System.Collections.Concurrent;
using System.IO;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using Microsoft.Data.Sqlite;
using TaskTool.Models;

namespace TaskTool.Services;

public sealed record AiRetrievalMatch(AiKnowledgeSourceType SourceType, string Content, string Title, string DisplaySource, string? Url, double Score, int? PageNumber = null,
    string? SpaceKey = null, string? SectionTitle = null, WikiKnowledgeContentKind? ContentKind = null, string? AttachmentName = null, string? ExternalPageId = null);
public enum AiKnowledgeSourceType { LocalFiles, Wiki }

public sealed class WikiAiKnowledgeService : IDisposable
{
    public static readonly TimeSpan SyncInterval = TimeSpan.FromMinutes(30);
    private readonly SettingsService _settings;
    private readonly LoggerService _logger;
    private readonly Dictionary<string, IWikiKnowledgeProvider> _providers;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private readonly ConfluenceKnowledgeParser _parser = new();
    private readonly WikiAttachmentProcessor _attachmentProcessor;
    private readonly Timer _timer;
    public string IndexPath { get; }
    public event EventHandler? StatusChanged;

    public WikiAiKnowledgeService(SettingsService settings, LoggerService logger, IEnumerable<(string Type, IWikiKnowledgeProvider Provider)>? providers = null, string? localAppData = null, WikiAttachmentProcessor? attachmentProcessor = null)
    {
        _settings = settings; _logger = logger; _attachmentProcessor = attachmentProcessor ?? new WikiAttachmentProcessor();
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
            var scopeChanged = !string.Equals(previousFingerprint, fingerprint, StringComparison.Ordinal);
            fullRebuild |= old.Count == 0 || scopeChanged;
            SetTransientStatus(source.Id, fullRebuild ? "full-index" : "syncing", scopeChanged);
            _logger.Info($"[Wiki AI Index] sourceId={source.Id} action=sync-start mode={(fullRebuild ? "full" : "incremental")}");
            var pages = new List<WikiKnowledgePage>(); const int size = 100; var offset = 0;
            while (true)
            {
                var batch = await _providers[source.ProviderType].GetPagesAsync(source, offset, size, token);
                pages.AddRange(batch.Pages.Where(page => WikiScopePolicy.AllowsSpace(source, page.SpaceKey))); offset += batch.Pages.Count;
                if (!batch.HasMore || batch.Pages.Count == 0) break;
            }
            var seen = pages.ToDictionary(x => x.ExternalId, StringComparer.Ordinal);
            var changedPages = new ConcurrentDictionary<string,WikiKnowledgePage>(pages.Where(p => fullRebuild || !old.TryGetValue(p.ExternalId, out var existing) || IsChanged(existing, p)).ToDictionary(p => p.ExternalId, StringComparer.Ordinal),StringComparer.Ordinal);
            var prefetchedAttachments = new Dictionary<string, IReadOnlyList<WikiKnowledgeAttachmentMetadata>>(StringComparer.Ordinal);
            if (!fullRebuild)
            {
                var pagesWithAttachments = LoadAttachmentPageIds(source.Id);
                using var metadataThrottle = new SemaphoreSlim(3, 3);
                await Task.WhenAll(pages.Where(page => pagesWithAttachments.Contains(page.ExternalId) && !changedPages.ContainsKey(page.ExternalId)).Select(async page =>
                {
                    await metadataThrottle.WaitAsync(token);
                    try
                    {
                        var remote = await _providers[source.ProviderType].GetAttachmentsAsync(source, page.ExternalId, token);
                        lock (prefetchedAttachments) prefetchedAttachments[page.ExternalId] = remote;
                        if (AttachmentMetadataChanged(source.Id, page.ExternalId, remote)) changedPages[page.ExternalId] = page;
                    }
                    finally { metadataThrottle.Release(); }
                }));
            }
            var changed = changedPages.Values.ToArray();
            var loaded = new Dictionary<string, WikiKnowledgePageContent>(StringComparer.Ordinal);
            var attachments = new Dictionary<string, IReadOnlyList<WikiKnowledgeAttachmentResult>>(StringComparer.Ordinal);
            var attachmentsDownloaded=0;var attachmentsReused=0;
            using var throttle = new SemaphoreSlim(3, 3);
            await Task.WhenAll(changed.Select(async page =>
            {
                await throttle.WaitAsync(token);
                try
                {
                    var content = await _providers[source.ProviderType].GetPageContentAsync(source, page.ExternalId, token);
                    var referenced = _parser.GetReferencedAttachmentNames(content.StorageMarkup); var processed = new List<WikiKnowledgeAttachmentResult>();
                    if (referenced.Count > 0)
                    {
                        var metadata = prefetchedAttachments.TryGetValue(page.ExternalId, out var prefetched) ? prefetched : await _providers[source.ProviderType].GetAttachmentsAsync(source, page.ExternalId, token);
                        foreach (var attachment in metadata.Where(item => referenced.Contains(item.FileName)))
                        {
                            var contextual = attachment with { SectionTitle = page.Title };
                            var cached = LoadCachedAttachment(source.Id, page.ExternalId, contextual);
                            if (cached != null) { Interlocked.Increment(ref attachmentsReused); processed.Add(cached); continue; }
                            try
                            {
                                Interlocked.Increment(ref attachmentsDownloaded);await using var stream = await _providers[source.ProviderType].DownloadAttachmentAsync(source, contextual, token); var payload=await _attachmentProcessor.ReadPayloadAsync(contextual,stream,token); var previous=LoadLatestAttachment(source.Id,page.ExternalId,contextual);
                                processed.Add(previous is { ContentHash: not null, ExtractionVersion: WikiAttachmentProcessor.CurrentExtractionVersion } && previous.ContentHash.Equals(payload.Sha256,StringComparison.OrdinalIgnoreCase)
                                    ? previous with { Metadata=contextual,Status="indexed" }
                                    : await _attachmentProcessor.ProcessPayloadAsync(contextual,payload,token));
                            }
                            catch (Exception ex) when (ex is not OperationCanceledException) { _logger.Warning($"[Wiki AI Attachment] sourceId={source.Id} pageId={page.ExternalId} attachmentId={attachment.AttachmentId} status=update-failed reason={AttachmentFailureReason(contextual, ex)}"); var stale=LoadLatestAttachment(source.Id,page.ExternalId,contextual); processed.Add(stale == null ? new(contextual,[new(page.Title,WikiKnowledgeContentKind.AttachmentMetadata,$"Attachment: {attachment.FileName}",0,attachment.AttachmentId,attachment.FileName)],"failed",null) : stale with { Metadata=contextual,Status="update-failed" }); }
                        }
                    }
                    lock (loaded)
                    {
                        loaded[page.ExternalId] = content;
                        attachments[page.ExternalId] = processed;
                    }
                }
                finally
                {
                    throttle.Release();
                }
            }));
            Apply(source, fingerprint, pages, loaded, attachments, fullRebuild);
            var deleted = old.Keys.Count(x => !seen.ContainsKey(x)); var added = changed.Count(x => !old.ContainsKey(x.ExternalId));
            _logger.Info($"[Wiki AI Index] sourceId={source.Id} pagesSeen={pages.Count} changed={changed.Length - added} new={added} deleted={deleted} unchanged={pages.Count - changed.Length}");
            var status = GetStatus(source.Id);
            _logger.Info($"[Wiki AI Index] sourceId={source.Id} action={(fullRebuild ? "full-index" : "sync-complete")} pages={status.PageCount} chunks={status.ChunkCount} attachments={status.AttachmentCount} attachmentsDownloaded={attachmentsDownloaded} attachmentsReused={attachmentsReused} pdfPages={status.PdfPageCount} ocrSuccess={status.OcrSuccessCount} drawio={status.DrawIoCount} durationMs={watch.ElapsedMilliseconds}");
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
        var eligible = _settings.Current.WikiSources.Where(IsIndexable)
            .Where(source => string.Equals(GetFingerprint(source.Id), WikiScopePolicy.Fingerprint(source), StringComparison.Ordinal))
            .ToDictionary(source => source.Id, StringComparer.Ordinal);
        if (eligible.Count == 0) return Array.Empty<AiRetrievalMatch>();
        var queryEvaluation = AiKnowledgeSearchService.EvaluateRelevance(terms, string.Empty, string.Empty, string.Empty);
        var watch = Stopwatch.StartNew(); var result = new List<(AiRetrievalMatch Match, RelevanceEvaluation Evaluation)>(); var candidates = 0; var rejectedAnchorMismatch = 0;
        await using var db = Open(true); await using var cmd = db.CreateCommand();
        var sourceParameters = eligible.Keys.Select((_, index) => $"$s{index}").ToArray();
        cmd.CommandText = $"SELECT c.content,c.title,s.name,c.url,c.space_key,c.source_id,c.external_id,c.section_title,c.content_kind,c.attachment_name,c.page_number,bm25(wiki_ai_chunks_fts,1,5,4,2,3) FROM wiki_ai_chunks_fts f JOIN wiki_ai_chunks c ON c.id=f.rowid JOIN wiki_ai_source_names s ON s.source_id=c.source_id WHERE wiki_ai_chunks_fts MATCH $q AND c.source_id IN ({string.Join(',', sourceParameters)}) LIMIT $l";
        cmd.Parameters.AddWithValue("$q", string.Join(" OR ", terms.Select(x => $"\"{x}\"*"))); cmd.Parameters.AddWithValue("$l", Math.Max(limit * 8, 32));
        var parameterIndex = 0; foreach (var sourceId in eligible.Keys) cmd.Parameters.AddWithValue(sourceParameters[parameterIndex++], sourceId);
        await using var reader = await cmd.ExecuteReaderAsync(token);
        while (await reader.ReadAsync(token))
        {
            candidates++; var source = eligible[reader.GetString(5)]; var space = reader.GetString(4);
            if (!WikiScopePolicy.AllowsSpace(source, space)) continue;
            var content = reader.GetString(0); var title = reader.GetString(1); var section = reader.GetString(7); var kind=Enum.Parse<WikiKnowledgeContentKind>(reader.GetString(8));
            var evaluation = AiKnowledgeSearchService.EvaluateRelevance(terms, content, title, section + " " + space);
            if (!evaluation.HasAnchorMatch) { rejectedAnchorMismatch++; continue; }
            var passesQualityGate = evaluation.IsRelevant && (evaluation.HasSpecificExactMatch || evaluation.HasStrongTitleOrSectionMatch
                || evaluation.MeaningfulCoverage >= .75 || evaluation.MeaningfulTermCount == 0);
            if (!passesQualityGate) continue;
            var score = evaluation.Score * ContentKindWeight(kind);
            result.Add((new(AiKnowledgeSourceType.Wiki, content, title, reader.GetString(2), reader.GetString(3), score,
                reader.IsDBNull(10) ? null : reader.GetInt32(10), space, section, kind, reader.IsDBNull(9) ? null : reader.GetString(9), reader.GetString(6)), evaluation));
        }
        var pages = result.GroupBy(x => $"{x.Match.DisplaySource}:{x.Match.ExternalPageId}", StringComparer.OrdinalIgnoreCase)
            .Select(group => { var ordered = group.OrderByDescending(x => x.Match.Score).ToArray(); var qualified = ordered.Where(x => x.Match.Score >= ordered[0].Match.Score * .70).Take(2).ToArray(); return new { Matches = qualified, Score = ordered[0].Match.Score + qualified.Skip(1).Select(x => x.Match.Score * .15).FirstOrDefault() }; })
            .OrderByDescending(page => page.Score).ToArray();
        var best = pages.FirstOrDefault()?.Score ?? 0; var minimumAcceptedScore = best * .72; var acceptedPages = pages.Where(page => page.Score >= minimumAcceptedScore).Take(3).ToArray();
        var acceptedEntries = acceptedPages.SelectMany(page => page.Matches).Take(limit).ToArray(); var accepted = acceptedEntries.Select(x => x.Match).ToArray();
        var bestCoverage = pages.Length == 0 || pages[0].Matches.Length == 0 ? 0 : pages[0].Matches[0].Evaluation.MeaningfulCoverage;
        var averageCoverage = acceptedEntries.Length == 0 ? 0 : acceptedEntries.Average(x => x.Evaluation.MeaningfulCoverage);
        _logger.OperationalInfo($"[AI Wiki] Query completed terms={terms.Length} anchors={queryEvaluation.AnchorCount} hardAnchors={queryEvaluation.HardAnchorCount} candidates={candidates} qualifiedCandidates={result.Count} rejectedAnchorMismatch={rejectedAnchorMismatch} acceptedPages={acceptedPages.Length} acceptedChunks={accepted.Length} bestScore={best:F2} minimumAcceptedScore={minimumAcceptedScore:F2} bestCoverage={bestCoverage:F2} acceptedAverageCoverage={averageCoverage:F2} durationMs={watch.ElapsedMilliseconds}"); return accepted;
    }

    public WikiAiIndexStatus GetStatus(string sourceId)
    {
        using var db = Open(); using var cmd = db.CreateCommand(); cmd.CommandText = "SELECT status,page_count,chunk_count,last_success_utc FROM wiki_ai_sources WHERE source_id=$s"; cmd.Parameters.AddWithValue("$s", sourceId);
        using var r = cmd.ExecuteReader(); if(!r.Read()) return new(sourceId,0,0,null,"not-indexed"); var basic=new WikiAiIndexStatus(sourceId,r.GetInt32(1),r.GetInt32(2),r.IsDBNull(3)?null:DateTime.Parse(r.GetString(3)).ToUniversalTime(),r.GetString(0));r.Close();using var counts=db.CreateCommand();counts.CommandText="SELECT count(*),sum(CASE WHEN lower(file_name) LIKE '%.pdf' THEN 1 ELSE 0 END),sum(pdf_page_count),sum(CASE WHEN media_type LIKE 'image/%' THEN 1 ELSE 0 END),sum(ocr_succeeded),sum(CASE WHEN media_type LIKE 'image/%' AND ocr_succeeded=0 AND status IN ('ocr-empty','ocr-unavailable','ocr-failed','update-failed','failed') THEN 1 ELSE 0 END),sum(CASE WHEN lower(file_name) LIKE '%.drawio' OR media_type LIKE '%drawio%' THEN 1 ELSE 0 END) FROM wiki_ai_attachments WHERE source_id=$s";counts.Parameters.AddWithValue("$s",sourceId);using var cr=counts.ExecuteReader();cr.Read();int V(int i)=>cr.IsDBNull(i)?0:cr.GetInt32(i);return basic with { AttachmentCount=V(0),PdfCount=V(1),PdfPageCount=V(2),ImageCount=V(3),OcrSuccessCount=V(4),OcrFailureCount=V(5),DrawIoCount=V(6) };
    }

    public void Invalidate(string sourceId) { using var db = Open(); using var cmd = db.CreateCommand(); cmd.CommandText = "UPDATE wiki_ai_sources SET scope_fingerprint='' WHERE source_id=$s"; cmd.Parameters.AddWithValue("$s", sourceId); cmd.ExecuteNonQuery(); }
    private Dictionary<string, (string Version, DateTime? Modified)> LoadPageMetadata(string id) { using var db = Open(); using var cmd = db.CreateCommand(); cmd.CommandText="SELECT external_id,version,last_modified_utc FROM wiki_ai_pages WHERE source_id=$s"; cmd.Parameters.AddWithValue("$s",id); using var r=cmd.ExecuteReader(); var d=new Dictionary<string,(string,DateTime?)>(); while(r.Read()) d[r.GetString(0)]=(r.GetString(1),r.IsDBNull(2)?null:DateTime.Parse(r.GetString(2)).ToUniversalTime()); return d; }
    private string? GetFingerprint(string id) { using var db=Open(); using var cmd=db.CreateCommand(); cmd.CommandText="SELECT scope_fingerprint FROM wiki_ai_sources WHERE source_id=$s";cmd.Parameters.AddWithValue("$s",id);return cmd.ExecuteScalar() as string; }
    private static bool IsChanged((string Version, DateTime? Modified) old, WikiKnowledgePage page) => (!string.IsNullOrEmpty(page.Version) && old.Version != page.Version) || (page.LastModifiedUtc.HasValue && old.Modified != page.LastModifiedUtc);
    private void Apply(WikiSourceSettings source,string fingerprint,IReadOnlyList<WikiKnowledgePage> pages,IReadOnlyDictionary<string,WikiKnowledgePageContent> loaded,IReadOnlyDictionary<string,IReadOnlyList<WikiKnowledgeAttachmentResult>> attachments,bool full)
    {
        using var db=Open(); using var tx=db.BeginTransaction();
        var ids=pages.Select(x=>x.ExternalId).ToHashSet(StringComparer.Ordinal); using(var q=db.CreateCommand()){q.Transaction=tx;q.CommandText="SELECT external_id FROM wiki_ai_pages WHERE source_id=$s";q.Parameters.AddWithValue("$s",source.Id);using var r=q.ExecuteReader();var deleted=new List<string>();while(r.Read())if(!ids.Contains(r.GetString(0)))deleted.Add(r.GetString(0));r.Close();foreach(var id in deleted)DeletePage(db,tx,source.Id,id);}
        foreach(var p in pages.Where(x=>loaded.ContainsKey(x.ExternalId))) { var attachmentResults=attachments.GetValueOrDefault(p.ExternalId)??[]; DeletePage(db,tx,source.Id,p.ExternalId); var body=loaded[p.ExternalId]; var parsed=_parser.Parse(p,body); var document=parsed with { Blocks=parsed.Blocks.Concat(attachmentResults.SelectMany(x=>x.Blocks)).ToArray() }; using var cmd=db.CreateCommand();cmd.Transaction=tx;cmd.CommandText="INSERT INTO wiki_ai_pages(source_id,external_id,title,url,space_key,version,last_modified_utc,content_hash,indexed_utc) VALUES($s,$e,$t,$u,$k,$v,$m,$h,$i)";cmd.Parameters.AddWithValue("$s",source.Id);cmd.Parameters.AddWithValue("$e",p.ExternalId);cmd.Parameters.AddWithValue("$t",p.Title);cmd.Parameters.AddWithValue("$u",p.Url);cmd.Parameters.AddWithValue("$k",p.SpaceKey);cmd.Parameters.AddWithValue("$v",body.Version);cmd.Parameters.AddWithValue("$m",body.LastModifiedUtc?.ToString("O")??(object)DBNull.Value);cmd.Parameters.AddWithValue("$h",Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(body.StorageMarkup??body.PlainText))));cmd.Parameters.AddWithValue("$i",DateTime.UtcNow.ToString("O"));cmd.ExecuteNonQuery(); foreach(var chunk in WikiKnowledgeChunker.Chunk(document)){InsertChunk(db,tx,p,chunk);} foreach(var attachment in attachmentResults){using var a=db.CreateCommand();a.Transaction=tx;a.CommandText="INSERT INTO wiki_ai_attachments(source_id,parent_external_id,attachment_id,file_name,media_type,version,last_modified_utc,content_hash,status,pdf_page_count,ocr_succeeded,extraction_version,indexed_utc) VALUES($s,$p,$a,$f,$m,$v,$l,$h,$status,$pages,$ocr,$extraction,$now)";a.Parameters.AddWithValue("$s",source.Id);a.Parameters.AddWithValue("$p",p.ExternalId);a.Parameters.AddWithValue("$a",attachment.Metadata.AttachmentId);a.Parameters.AddWithValue("$f",attachment.Metadata.FileName);a.Parameters.AddWithValue("$m",attachment.Metadata.MediaType);a.Parameters.AddWithValue("$v",attachment.Metadata.Version);a.Parameters.AddWithValue("$l",attachment.Metadata.LastModifiedUtc?.ToString("O")??(object)DBNull.Value);a.Parameters.AddWithValue("$h",(object?)attachment.ContentHash??DBNull.Value);a.Parameters.AddWithValue("$status",attachment.Status);a.Parameters.AddWithValue("$pages",attachment.PdfPageCount);a.Parameters.AddWithValue("$ocr",attachment.OcrSucceeded?1:0);a.Parameters.AddWithValue("$extraction",attachment.ExtractionVersion);a.Parameters.AddWithValue("$now",DateTime.UtcNow.ToString("O"));a.ExecuteNonQuery();}}
        using(var n=db.CreateCommand()){n.Transaction=tx;n.CommandText="INSERT OR REPLACE INTO wiki_ai_source_names(source_id,name) VALUES($s,$n)";n.Parameters.AddWithValue("$s",source.Id);n.Parameters.AddWithValue("$n",source.Name);n.ExecuteNonQuery();}
        using(var state=db.CreateCommand()){state.Transaction=tx;state.CommandText="INSERT INTO wiki_ai_sources(source_id,scope_fingerprint,last_full_sync_utc,last_incremental_sync_utc,last_success_utc,status,page_count,chunk_count) VALUES($s,$f,$full,$inc,$now,'current',(SELECT count(*) FROM wiki_ai_pages WHERE source_id=$s),(SELECT count(*) FROM wiki_ai_chunks WHERE source_id=$s)) ON CONFLICT(source_id) DO UPDATE SET scope_fingerprint=$f,last_full_sync_utc=COALESCE($full,last_full_sync_utc),last_incremental_sync_utc=$inc,last_success_utc=$now,status='current',page_count=(SELECT count(*) FROM wiki_ai_pages WHERE source_id=$s),chunk_count=(SELECT count(*) FROM wiki_ai_chunks WHERE source_id=$s)";state.Parameters.AddWithValue("$s",source.Id);state.Parameters.AddWithValue("$f",fingerprint);state.Parameters.AddWithValue("$full",full?DateTime.UtcNow.ToString("O"):(object)DBNull.Value);state.Parameters.AddWithValue("$inc",DateTime.UtcNow.ToString("O"));state.Parameters.AddWithValue("$now",DateTime.UtcNow.ToString("O"));state.ExecuteNonQuery();} tx.Commit();
    }
    private static void InsertChunk(SqliteConnection db,SqliteTransaction tx,WikiKnowledgePage p,WikiKnowledgeBlock chunk){using var c=db.CreateCommand();c.Transaction=tx;c.CommandText="INSERT INTO wiki_ai_chunks(page_id,chunk_index,content,title,space_key,source_id,url,external_id,section_title,content_kind,attachment_id,attachment_name,page_number,block_ordinal) VALUES((SELECT id FROM wiki_ai_pages WHERE source_id=$s AND external_id=$e),$n,$c,$t,$k,$s,$u,$e,$section,$kind,$attachmentId,$attachmentName,$page,$ordinal)";c.Parameters.AddWithValue("$s",p.SourceId);c.Parameters.AddWithValue("$e",p.ExternalId);c.Parameters.AddWithValue("$n",chunk.BlockOrdinal);c.Parameters.AddWithValue("$c",chunk.Content);c.Parameters.AddWithValue("$t",p.Title);c.Parameters.AddWithValue("$k",p.SpaceKey);c.Parameters.AddWithValue("$u",p.Url);c.Parameters.AddWithValue("$section",chunk.SectionTitle);c.Parameters.AddWithValue("$kind",chunk.ContentKind.ToString());c.Parameters.AddWithValue("$attachmentId",(object?)chunk.AttachmentId??DBNull.Value);c.Parameters.AddWithValue("$attachmentName",(object?)chunk.AttachmentName??DBNull.Value);c.Parameters.AddWithValue("$page",(object?)chunk.PageNumber??DBNull.Value);c.Parameters.AddWithValue("$ordinal",chunk.BlockOrdinal);c.ExecuteNonQuery();}
    private static void DeletePage(SqliteConnection db,SqliteTransaction tx,string source,string external){using var c=db.CreateCommand();c.Transaction=tx;c.CommandText="DELETE FROM wiki_ai_chunks WHERE source_id=$s AND external_id=$e; DELETE FROM wiki_ai_attachments WHERE source_id=$s AND parent_external_id=$e; DELETE FROM wiki_ai_pages WHERE source_id=$s AND external_id=$e";c.Parameters.AddWithValue("$s",source);c.Parameters.AddWithValue("$e",external);c.ExecuteNonQuery();}
    private WikiKnowledgeAttachmentResult? LoadCachedAttachment(string source,string page,WikiKnowledgeAttachmentMetadata metadata){using var db=Open();using var c=db.CreateCommand();c.CommandText="SELECT content_hash,status,pdf_page_count,ocr_succeeded,extraction_version FROM wiki_ai_attachments WHERE source_id=$s AND parent_external_id=$p AND attachment_id=$a AND version=$v AND COALESCE(last_modified_utc,'')=COALESCE($m,'')";c.Parameters.AddWithValue("$s",source);c.Parameters.AddWithValue("$p",page);c.Parameters.AddWithValue("$a",metadata.AttachmentId);c.Parameters.AddWithValue("$v",metadata.Version);c.Parameters.AddWithValue("$m",metadata.LastModifiedUtc?.ToString("O")??(object)DBNull.Value);using var r=c.ExecuteReader();if(!r.Read())return null;var hash=r.IsDBNull(0)?null:r.GetString(0);var status=r.GetString(1);var pages=r.GetInt32(2);var ocr=r.GetInt32(3)!=0;var extraction=r.GetInt32(4);r.Close();using var chunks=db.CreateCommand();chunks.CommandText="SELECT section_title,content_kind,content,block_ordinal,page_number FROM wiki_ai_chunks WHERE source_id=$s AND external_id=$p AND attachment_id=$a ORDER BY block_ordinal,page_number";chunks.Parameters.AddWithValue("$s",source);chunks.Parameters.AddWithValue("$p",page);chunks.Parameters.AddWithValue("$a",metadata.AttachmentId);using var reader=chunks.ExecuteReader();var blocks=new List<WikiKnowledgeBlock>();while(reader.Read())blocks.Add(new(reader.GetString(0),Enum.Parse<WikiKnowledgeContentKind>(reader.GetString(1)),reader.GetString(2),reader.GetInt32(3),metadata.AttachmentId,metadata.FileName,reader.IsDBNull(4)?null:reader.GetInt32(4)));return new(metadata,blocks,status,hash,pages,ocr,extraction);}
    private WikiKnowledgeAttachmentResult? LoadLatestAttachment(string source,string page,WikiKnowledgeAttachmentMetadata metadata){using var db=Open();using var c=db.CreateCommand();c.CommandText="SELECT version,last_modified_utc FROM wiki_ai_attachments WHERE source_id=$s AND parent_external_id=$p AND attachment_id=$a";c.Parameters.AddWithValue("$s",source);c.Parameters.AddWithValue("$p",page);c.Parameters.AddWithValue("$a",metadata.AttachmentId);using var r=c.ExecuteReader();if(!r.Read())return null;var cached=metadata with { Version=r.GetString(0),LastModifiedUtc=r.IsDBNull(1)?null:DateTime.Parse(r.GetString(1)).ToUniversalTime() };return LoadCachedAttachment(source,page,cached);}
    private HashSet<string> LoadAttachmentPageIds(string source){using var db=Open();using var c=db.CreateCommand();c.CommandText="SELECT DISTINCT parent_external_id FROM wiki_ai_attachments WHERE source_id=$s";c.Parameters.AddWithValue("$s",source);using var r=c.ExecuteReader();var result=new HashSet<string>(StringComparer.Ordinal);while(r.Read())result.Add(r.GetString(0));return result;}
    private bool AttachmentMetadataChanged(string source,string page,IReadOnlyList<WikiKnowledgeAttachmentMetadata> remote){using var db=Open();using var c=db.CreateCommand();c.CommandText="SELECT attachment_id,file_name,media_type,version,last_modified_utc FROM wiki_ai_attachments WHERE source_id=$s AND parent_external_id=$p";c.Parameters.AddWithValue("$s",source);c.Parameters.AddWithValue("$p",page);using var r=c.ExecuteReader();var local=new Dictionary<string,(string Name,string Media,string Version,string Modified)>(StringComparer.Ordinal);while(r.Read())local[r.GetString(0)]=(r.GetString(1),r.GetString(2),r.GetString(3),r.IsDBNull(4)?string.Empty:r.GetString(4));var relevantRemote=remote.Where(item=>local.ContainsKey(item.AttachmentId)).ToArray();if(relevantRemote.Length!=local.Count)return true;return relevantRemote.Any(item=>{var old=local[item.AttachmentId];return old.Name!=item.FileName||old.Media!=item.MediaType||old.Version!=item.Version||old.Modified!=(item.LastModifiedUtc?.ToString("O")??string.Empty);});}
    private static double ContentKindWeight(WikiKnowledgeContentKind kind)=>kind switch{WikiKnowledgeContentKind.AttachmentMetadata=>.55,WikiKnowledgeContentKind.ImageOcr=>.9,WikiKnowledgeContentKind.Table or WikiKnowledgeContentKind.DrawIo=>1.1,_=>1};
    private static string AttachmentFailureReason(WikiKnowledgeAttachmentMetadata attachment, Exception error)
    {
        if (error is HttpRequestException) return "download-failed";
        if (error is InvalidDataException && error.Message.Contains("size limit", StringComparison.OrdinalIgnoreCase)) return "too-large";
        var extension = Path.GetExtension(attachment.FileName);
        if (extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase) || attachment.MediaType.Equals("application/pdf", StringComparison.OrdinalIgnoreCase)) return "invalid-pdf";
        if (extension.Equals(".drawio", StringComparison.OrdinalIgnoreCase) || attachment.MediaType.Contains("drawio", StringComparison.OrdinalIgnoreCase)) return "invalid-drawio";
        if (attachment.MediaType.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) return "invalid-image";
        return "unsupported-format";
    }
    private void SetTransientStatus(string id,string status,bool invalidateFingerprint){using var db=Open();using var c=db.CreateCommand();c.CommandText="INSERT INTO wiki_ai_sources(source_id,scope_fingerprint,status,page_count,chunk_count) VALUES($s,'',$x,0,0) ON CONFLICT(source_id) DO UPDATE SET status=$x,scope_fingerprint=CASE WHEN $invalidate=1 THEN '' ELSE scope_fingerprint END";c.Parameters.AddWithValue("$s",id);c.Parameters.AddWithValue("$x",status);c.Parameters.AddWithValue("$invalidate",invalidateFingerprint?1:0);c.ExecuteNonQuery();StatusChanged?.Invoke(this,EventArgs.Empty);}
    private void MarkFailed(string id){using var db=Open();using var c=db.CreateCommand();c.CommandText="INSERT INTO wiki_ai_sources(source_id,scope_fingerprint,status,page_count,chunk_count) VALUES($s,'','failed',0,0) ON CONFLICT(source_id) DO UPDATE SET status='failed'";c.Parameters.AddWithValue("$s",id);c.ExecuteNonQuery();}
    private SqliteConnection Open(bool readOnly=false){var db=new SqliteConnection($"Data Source={IndexPath}{(readOnly?";Mode=ReadOnly":"")}");db.Open();using var pragma=db.CreateCommand();pragma.CommandText="PRAGMA busy_timeout=5000";pragma.ExecuteNonQuery();return db;}
    private void Initialize(){using var db=Open();using(var journal=db.CreateCommand()){journal.CommandText="PRAGMA journal_mode=WAL; PRAGMA synchronous=NORMAL;";journal.ExecuteNonQuery();}using(var cmd=db.CreateCommand()){cmd.CommandText="""
CREATE TABLE IF NOT EXISTS wiki_ai_sources(source_id TEXT PRIMARY KEY,scope_fingerprint TEXT NOT NULL,last_full_sync_utc TEXT,last_incremental_sync_utc TEXT,last_success_utc TEXT,status TEXT NOT NULL,page_count INTEGER NOT NULL,chunk_count INTEGER NOT NULL);
CREATE TABLE IF NOT EXISTS wiki_ai_source_names(source_id TEXT PRIMARY KEY,name TEXT NOT NULL);
CREATE TABLE IF NOT EXISTS wiki_ai_pages(id INTEGER PRIMARY KEY,source_id TEXT NOT NULL,external_id TEXT NOT NULL,title TEXT NOT NULL,url TEXT NOT NULL,space_key TEXT NOT NULL,version TEXT NOT NULL,last_modified_utc TEXT,content_hash TEXT NOT NULL,indexed_utc TEXT NOT NULL,UNIQUE(source_id,external_id));
CREATE TABLE IF NOT EXISTS wiki_ai_attachments(source_id TEXT NOT NULL,parent_external_id TEXT NOT NULL,attachment_id TEXT NOT NULL,file_name TEXT NOT NULL,media_type TEXT NOT NULL,version TEXT NOT NULL,last_modified_utc TEXT,content_hash TEXT,status TEXT NOT NULL,pdf_page_count INTEGER NOT NULL DEFAULT 0,ocr_succeeded INTEGER NOT NULL DEFAULT 0,extraction_version INTEGER NOT NULL DEFAULT 1,indexed_utc TEXT NOT NULL,PRIMARY KEY(source_id,parent_external_id,attachment_id));
CREATE TABLE IF NOT EXISTS wiki_ai_chunks(id INTEGER PRIMARY KEY,page_id INTEGER NOT NULL,chunk_index INTEGER NOT NULL,content TEXT NOT NULL,title TEXT NOT NULL,space_key TEXT NOT NULL,source_id TEXT NOT NULL,url TEXT NOT NULL,external_id TEXT NOT NULL,FOREIGN KEY(page_id) REFERENCES wiki_ai_pages(id) ON DELETE CASCADE);
""";cmd.ExecuteNonQuery();}
        EnsureColumn(db,"wiki_ai_attachments","extraction_version","INTEGER NOT NULL DEFAULT 1");
        var migrated = EnsureColumn(db,"wiki_ai_chunks","section_title","TEXT NOT NULL DEFAULT ''") | EnsureColumn(db,"wiki_ai_chunks","content_kind","TEXT NOT NULL DEFAULT 'PageText'") | EnsureColumn(db,"wiki_ai_chunks","attachment_id","TEXT") | EnsureColumn(db,"wiki_ai_chunks","attachment_name","TEXT") | EnsureColumn(db,"wiki_ai_chunks","page_number","INTEGER") | EnsureColumn(db,"wiki_ai_chunks","block_ordinal","INTEGER NOT NULL DEFAULT 0");
        using(var indexes=db.CreateCommand()){indexes.CommandText="CREATE INDEX IF NOT EXISTS ix_wiki_ai_attachments_source_parent ON wiki_ai_attachments(source_id,parent_external_id); CREATE INDEX IF NOT EXISTS ix_wiki_ai_chunks_source_page_attachment ON wiki_ai_chunks(source_id,external_id,attachment_id);";indexes.ExecuteNonQuery();}
        using var inspect=db.CreateCommand();inspect.CommandText="SELECT sql FROM sqlite_master WHERE type='table' AND name='wiki_ai_chunks_fts'";var ftsSql=inspect.ExecuteScalar() as string;
        var rebuildFts=migrated || ftsSql==null || !ftsSql.Contains("section_title",StringComparison.OrdinalIgnoreCase);
        if(rebuildFts){using var drop=db.CreateCommand();drop.CommandText="DROP TRIGGER IF EXISTS wiki_ai_chunks_ai; DROP TRIGGER IF EXISTS wiki_ai_chunks_ad; DROP TABLE IF EXISTS wiki_ai_chunks_fts;";drop.ExecuteNonQuery();}
        using var fts=db.CreateCommand();fts.CommandText="""
CREATE VIRTUAL TABLE IF NOT EXISTS wiki_ai_chunks_fts USING fts5(content,title,section_title,space_key,attachment_name,content='wiki_ai_chunks',content_rowid='id');
CREATE TRIGGER IF NOT EXISTS wiki_ai_chunks_ai AFTER INSERT ON wiki_ai_chunks BEGIN INSERT INTO wiki_ai_chunks_fts(rowid,content,title,section_title,space_key,attachment_name) VALUES(new.id,new.content,new.title,new.section_title,new.space_key,new.attachment_name); END;
CREATE TRIGGER IF NOT EXISTS wiki_ai_chunks_ad AFTER DELETE ON wiki_ai_chunks BEGIN INSERT INTO wiki_ai_chunks_fts(wiki_ai_chunks_fts,rowid,content,title,section_title,space_key,attachment_name) VALUES('delete',old.id,old.content,old.title,old.section_title,old.space_key,old.attachment_name); END;
""";fts.ExecuteNonQuery();if(rebuildFts){using var rebuild=db.CreateCommand();rebuild.CommandText="INSERT INTO wiki_ai_chunks_fts(wiki_ai_chunks_fts) VALUES('rebuild')";rebuild.ExecuteNonQuery();}}
    private static bool EnsureColumn(SqliteConnection db,string table,string name,string definition){using var check=db.CreateCommand();check.CommandText=$"SELECT count(*) FROM pragma_table_info('{table}') WHERE name=$name";check.Parameters.AddWithValue("$name",name);if(Convert.ToInt32(check.ExecuteScalar())>0)return false;using var alter=db.CreateCommand();alter.CommandText=$"ALTER TABLE {table} ADD COLUMN {name} {definition}";alter.ExecuteNonQuery();return true;}
    public void Dispose(){_timer.Dispose();_gate.Dispose();}
}
