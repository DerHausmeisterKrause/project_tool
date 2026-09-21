using System.Security.Cryptography;
using Microsoft.Data.Sqlite;

namespace TaskTool.Services;

public sealed record AiKnowledgeIndexStatus(int DocumentCount, int ChunkCount, DateTime? LastIndexedUtc, string? Error = null);

public sealed class AiKnowledgeIndexService
{
    private readonly IAiKnowledgeDocumentExtractor _extractor;
    private readonly LoggerService _logger;
    private readonly SemaphoreSlim _gate = new(1, 1);
    public string KnowledgePath { get; }
    public string IndexPath { get; }

    public AiKnowledgeIndexService(LoggerService logger, IAiKnowledgeDocumentExtractor? extractor = null, string? localAppData = null)
    {
        _logger = logger;
        _extractor = extractor ?? new AiKnowledgeDocumentExtractor();
        var root = Path.Combine(localAppData ?? Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Plenaro", "AI");
        KnowledgePath = Path.Combine(root, "knowledge");
        IndexPath = Path.Combine(root, "knowledge-index.db");
    }

    public void EnsureKnowledgeDirectory() => Directory.CreateDirectory(KnowledgePath);

    public async Task<AiKnowledgeIndexStatus> IndexAsync(bool rebuild = false, CancellationToken cancellationToken = default)
    {
        await _gate.WaitAsync(cancellationToken);
        try
        {
            EnsureKnowledgeDirectory();
            Directory.CreateDirectory(Path.GetDirectoryName(IndexPath)!);
            await using var connection = new SqliteConnection($"Data Source={IndexPath}");
            await connection.OpenAsync(cancellationToken);
            await InitializeSchemaAsync(connection, cancellationToken);
            if (rebuild) await ExecuteAsync(connection, "DELETE FROM knowledge_documents", cancellationToken);

            var files = Directory.EnumerateFiles(KnowledgePath, "*", SearchOption.AllDirectories)
                .Where(_extractor.IsSupported).ToArray();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var path in files)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var relative = Path.GetRelativePath(KnowledgePath, path).Replace('/', '\\');
                seen.Add(relative);
                var info = new FileInfo(path);
                if (!rebuild && await IsUnchangedAsync(connection, relative, info, cancellationToken)) continue;
                await ReplaceDocumentAsync(connection, path, relative, info, cancellationToken);
            }
            await DeleteMissingAsync(connection, seen, cancellationToken);
            return await GetStatusAsync(connection, cancellationToken);
        }
        finally { _gate.Release(); }
    }

    private async Task ReplaceDocumentAsync(SqliteConnection db, string path, string relative, FileInfo info, CancellationToken ct)
    {
        try
        {
            var extracted = await _extractor.ExtractAsync(path, ct);
            var chunks = AiKnowledgeChunker.Chunk(extracted);
            var status = chunks.Count == 0 && Path.GetExtension(path).Equals(".pdf", StringComparison.OrdinalIgnoreCase)
                ? "PDF enthält keinen extrahierbaren Text." : "Indexed";
            await DeleteDocumentAsync(db, relative, ct);
            await using var insert = db.CreateCommand();
            insert.CommandText = "INSERT INTO knowledge_documents(relative_path,category_path,file_name,extension,file_size,last_write_utc,content_hash,indexed_utc,status) VALUES($p,$c,$n,$e,$s,$w,$h,$i,$status); SELECT last_insert_rowid();";
            insert.Parameters.AddWithValue("$p", relative);
            insert.Parameters.AddWithValue("$c", Path.GetDirectoryName(relative) ?? string.Empty);
            insert.Parameters.AddWithValue("$n", Path.GetFileName(relative));
            insert.Parameters.AddWithValue("$e", Path.GetExtension(path).ToLowerInvariant());
            insert.Parameters.AddWithValue("$s", info.Length);
            insert.Parameters.AddWithValue("$w", info.LastWriteTimeUtc.ToString("O"));
            insert.Parameters.AddWithValue("$h", Convert.ToHexString(SHA256.HashData(await File.ReadAllBytesAsync(path, ct))));
            insert.Parameters.AddWithValue("$i", DateTime.UtcNow.ToString("O"));
            insert.Parameters.AddWithValue("$status", status);
            var id = (long)(await insert.ExecuteScalarAsync(ct) ?? 0L);
            for (var index = 0; index < chunks.Count; index++)
            {
                await using var chunk = db.CreateCommand();
                chunk.CommandText = "INSERT INTO knowledge_chunks(document_id,chunk_index,content,category_path,file_name,relative_path,page_number) VALUES($d,$i,$t,$c,$n,$r,$page)";
                chunk.Parameters.AddWithValue("$d", id); chunk.Parameters.AddWithValue("$i", index); chunk.Parameters.AddWithValue("$t", chunks[index].Content);
                chunk.Parameters.AddWithValue("$c", Path.GetDirectoryName(relative) ?? string.Empty); chunk.Parameters.AddWithValue("$n", Path.GetFileName(relative)); chunk.Parameters.AddWithValue("$r", relative);
                chunk.Parameters.AddWithValue("$page", (object?)chunks[index].PageNumber ?? DBNull.Value);
                await chunk.ExecuteNonQueryAsync(ct);
            }
        }
        catch (Exception ex) { _logger.Warning($"[AI Knowledge] Index failed file='{relative}' error='{ex.Message}'"); }
    }

    private static async Task InitializeSchemaAsync(SqliteConnection db, CancellationToken ct)
    {
        const string sql = """
            PRAGMA foreign_keys=ON;
            CREATE TABLE IF NOT EXISTS knowledge_documents(id INTEGER PRIMARY KEY,relative_path TEXT NOT NULL UNIQUE,category_path TEXT NOT NULL,file_name TEXT NOT NULL,extension TEXT NOT NULL,file_size INTEGER NOT NULL,last_write_utc TEXT NOT NULL,content_hash TEXT NOT NULL,indexed_utc TEXT NOT NULL,status TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS knowledge_chunks(id INTEGER PRIMARY KEY,document_id INTEGER NOT NULL REFERENCES knowledge_documents(id) ON DELETE CASCADE,chunk_index INTEGER NOT NULL,content TEXT NOT NULL,category_path TEXT NOT NULL,file_name TEXT NOT NULL,relative_path TEXT NOT NULL,page_number INTEGER);
            CREATE VIRTUAL TABLE IF NOT EXISTS knowledge_chunks_fts USING fts5(content,file_name,category_path,content='knowledge_chunks',content_rowid='id');
            CREATE TRIGGER IF NOT EXISTS chunks_ai AFTER INSERT ON knowledge_chunks BEGIN INSERT INTO knowledge_chunks_fts(rowid,content,file_name,category_path) VALUES(new.id,new.content,new.file_name,new.category_path); END;
            CREATE TRIGGER IF NOT EXISTS chunks_ad AFTER DELETE ON knowledge_chunks BEGIN INSERT INTO knowledge_chunks_fts(knowledge_chunks_fts,rowid,content,file_name,category_path) VALUES('delete',old.id,old.content,old.file_name,old.category_path); END;
            """;
        await ExecuteAsync(db, sql, ct);
    }

    private static async Task<bool> IsUnchangedAsync(SqliteConnection db, string relative, FileInfo info, CancellationToken ct)
    {
        await using var command = db.CreateCommand(); command.CommandText = "SELECT COUNT(*) FROM knowledge_documents WHERE relative_path=$p AND file_size=$s AND last_write_utc=$w";
        command.Parameters.AddWithValue("$p", relative); command.Parameters.AddWithValue("$s", info.Length); command.Parameters.AddWithValue("$w", info.LastWriteTimeUtc.ToString("O"));
        return Convert.ToInt32(await command.ExecuteScalarAsync(ct)) > 0;
    }
    private static async Task DeleteMissingAsync(SqliteConnection db, HashSet<string> seen, CancellationToken ct)
    {
        await using var cmd = db.CreateCommand(); cmd.CommandText = "SELECT relative_path FROM knowledge_documents";
        var existing = new List<string>(); await using (var reader = await cmd.ExecuteReaderAsync(ct)) while (await reader.ReadAsync(ct)) existing.Add(reader.GetString(0));
        foreach (var path in existing.Where(path => !seen.Contains(path))) await DeleteDocumentAsync(db, path, ct);
    }
    private static async Task DeleteDocumentAsync(SqliteConnection db, string relative, CancellationToken ct) { await using var cmd = db.CreateCommand(); cmd.CommandText = "DELETE FROM knowledge_documents WHERE relative_path=$p"; cmd.Parameters.AddWithValue("$p", relative); await cmd.ExecuteNonQueryAsync(ct); }
    private static async Task ExecuteAsync(SqliteConnection db, string sql, CancellationToken ct) { await using var cmd = db.CreateCommand(); cmd.CommandText = sql; await cmd.ExecuteNonQueryAsync(ct); }
    public async Task<AiKnowledgeIndexStatus> GetStatusAsync(CancellationToken ct = default) { await using var db = new SqliteConnection($"Data Source={IndexPath}"); await db.OpenAsync(ct); await InitializeSchemaAsync(db, ct); return await GetStatusAsync(db, ct); }
    private static async Task<AiKnowledgeIndexStatus> GetStatusAsync(SqliteConnection db, CancellationToken ct)
    {
        await using var cmd = db.CreateCommand(); cmd.CommandText = "SELECT (SELECT COUNT(*) FROM knowledge_documents WHERE status='Indexed'),(SELECT COUNT(*) FROM knowledge_chunks),MAX(indexed_utc) FROM knowledge_documents";
        await using var reader = await cmd.ExecuteReaderAsync(ct); await reader.ReadAsync(ct);
        return new(reader.GetInt32(0), reader.GetInt32(1), reader.IsDBNull(2) ? null : DateTime.Parse(reader.GetString(2)).ToUniversalTime());
    }
}
