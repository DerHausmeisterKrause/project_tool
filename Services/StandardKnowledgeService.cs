using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace TaskTool.Services;

public enum StandardKnowledgeState { NotInstalled, Checking, Downloading, Verifying, Installing, Indexing, Installed, DownloadFailed, Sha256Failed }
public sealed record StandardKnowledgeInstallation(string ReleaseTag, string ArchiveFile, string Sha256, DateTime InstalledUtc);
public sealed record StandardKnowledgeStatus(StandardKnowledgeState State, int DocumentCount, string? ReleaseTag, string? Error = null);

public sealed class StandardKnowledgeService
{
    public const string ArchiveName = "plenaro_it_wissensbasis.zip";
    private static readonly string[] SupportedExtensions = [".md", ".txt", ".pdf"];
    private static readonly JsonSerializerOptions JsonOptions = new(JsonSerializerDefaults.Web);
    private readonly HttpClient _http;
    private readonly LoggerService _logger;
    private readonly Func<string> _releaseTag;
    private readonly Func<Task> _reindex;
    private readonly SemaphoreSlim _gate = new(1, 1);
    public string RootPath { get; }
    public string KnowledgePath { get; }
    public string MetadataPath { get; }
    public event EventHandler? StatusChanged;
    public StandardKnowledgeStatus Status { get; private set; } = new(StandardKnowledgeState.NotInstalled, 0, null);

    public StandardKnowledgeService(AppVersionService version, AiKnowledgeIndexService index, LoggerService logger, HttpClient? http = null)
        : this(Path.GetDirectoryName(index.IndexPath)!, () => "V" + version.InstalledVersionText, async () => { await index.IndexAsync(); }, logger, http) { }

    internal StandardKnowledgeService(string root, Func<string> releaseTag, Func<Task> reindex, LoggerService logger, HttpClient? http = null)
    {
        RootPath = root; KnowledgePath = Path.Combine(root, "knowledge-default"); MetadataPath = Path.Combine(root, "standard-knowledge");
        _releaseTag = releaseTag; _reindex = reindex; _logger = logger; _http = http ?? new HttpClient();
        RefreshStatus();
    }

    public StandardKnowledgeStatus GetStatus() => Status;
    public async Task ForceCheckAsync(CancellationToken ct = default) => await EnsureInstalledAsync(ct);

    /// <summary>Ensures that the release knowledge is present and returns whether installation also reindexed it.</summary>
    public async Task<bool> EnsureInstalledAsync(CancellationToken ct = default)
    {
        await _gate.WaitAsync(ct);
        try
        {
            var tag = _releaseTag();
            SetStatus(StandardKnowledgeState.Checking);
            _logger.Info($"[Standard Knowledge] action=check release={tag}");
            var installed = await ReadInstalledAsync(ct);
            var expected = await DownloadExpectedShaAsync(tag, ct);
            _logger.Info($"[Standard Knowledge] remoteSha={expected} localSha={installed?.Sha256 ?? "none"}");
            if (installed != null && expected.Equals(installed.Sha256, StringComparison.OrdinalIgnoreCase) && Directory.Exists(KnowledgePath))
            {
                if (!tag.Equals(installed.ReleaseTag, StringComparison.Ordinal))
                    await WriteInstalledAsync(installed with { ReleaseTag = tag, InstalledUtc = DateTime.UtcNow }, ct);
                RefreshStatus(); return false;
            }

            var downloads = Path.Combine(RootPath, "downloads"); Directory.CreateDirectory(downloads);
            var part = Path.Combine(downloads, ArchiveName + ".part");
            try
            {
                SetStatus(StandardKnowledgeState.Downloading); _logger.Info("[Standard Knowledge] action=download-start");
                await DownloadAsync(AssetUri(tag, ArchiveName), part, ct);
                _logger.Info($"[Standard Knowledge] action=download-complete bytes={new FileInfo(part).Length}");
                SetStatus(StandardKnowledgeState.Verifying);
                var actual = await CalculateShaAsync(part, ct);
                if (!actual.Equals(expected, StringComparison.OrdinalIgnoreCase))
                {
                    _logger.Warning($"[Standard Knowledge] sha256Verified=false expected={expected} actual={actual}");
                    SetStatus(StandardKnowledgeState.Sha256Failed, "SHA256-Prüfung fehlgeschlagen"); return false;
                }
                _logger.Info("[Standard Knowledge] sha256Verified=true");
                SetStatus(StandardKnowledgeState.Installing);
                await ExtractSafelyAsync(part, tag, expected, ct);
                SetStatus(StandardKnowledgeState.Indexing); await _reindex();
                RefreshStatus(); _logger.Info($"[Standard Knowledge] action=install-complete documents={Status.DocumentCount}");
                return true;
            }
            finally { if (File.Exists(part)) File.Delete(part); }
        }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex)
        {
            _logger.Warning($"[Standard Knowledge] action=check-failed error='{ex.Message}'");
            RefreshStatus(ex.Message);
            return false;
        }
        finally { _gate.Release(); }
    }

    internal async Task ExtractSafelyAsync(string archive, string releaseTag, string sha, CancellationToken ct = default)
    {
        var staging = KnowledgePath + ".new"; var old = KnowledgePath + ".old";
        var metadataNew = MetadataPath + ".new"; var metadataOld = MetadataPath + ".old";
        DeleteDirectory(staging); DeleteDirectory(metadataNew); DeleteDirectory(old); DeleteDirectory(metadataOld);
        Directory.CreateDirectory(staging); Directory.CreateDirectory(metadataNew);
        try
        {
            using var zip = ZipFile.OpenRead(archive);
            var knowledgeEntries = zip.Entries.Where(e => EntryParts(e.FullName) is { Length: > 1 } p && p[0].Equals("knowledge", StringComparison.OrdinalIgnoreCase)).ToArray();
            if (!knowledgeEntries.Any(e => SupportedExtensions.Contains(Path.GetExtension(e.Name), StringComparer.OrdinalIgnoreCase))) throw new InvalidDataException("Das Archiv enthält keine unterstützten Knowledge-Dateien.");
            foreach (var entry in zip.Entries)
            {
                ct.ThrowIfCancellationRequested(); var parts = EntryParts(entry.FullName);
                if (parts.Length == 0) continue;
                string? destination = null;
                if (parts.Length > 1 && parts[0].Equals("knowledge", StringComparison.OrdinalIgnoreCase)) destination = SafePath(staging, parts[1..]);
                else if (parts.Length == 1 && (parts[0].Equals("MANIFEST.json", StringComparison.OrdinalIgnoreCase) || parts[0].Equals("STRUCTURE.txt", StringComparison.OrdinalIgnoreCase))) destination = SafePath(metadataNew, parts);
                if (destination == null || string.IsNullOrEmpty(entry.Name)) continue;
                Directory.CreateDirectory(Path.GetDirectoryName(destination)!); entry.ExtractToFile(destination, true);
            }
            await WriteInstalledAsync(new(releaseTag, ArchiveName, sha, DateTime.UtcNow), ct, metadataNew);
            var knowledgeBackedUp = false; var metadataBackedUp = false;
            var knowledgeActivated = false; var metadataActivated = false;
            try
            {
                if (Directory.Exists(KnowledgePath)) { Directory.Move(KnowledgePath, old); knowledgeBackedUp = true; }
                if (Directory.Exists(MetadataPath)) { Directory.Move(MetadataPath, metadataOld); metadataBackedUp = true; }
                Directory.Move(staging, KnowledgePath); knowledgeActivated = true;
                Directory.Move(metadataNew, MetadataPath); metadataActivated = true;
            }
            catch
            {
                if (metadataActivated) DeleteDirectory(MetadataPath);
                if (knowledgeActivated) DeleteDirectory(KnowledgePath);
                if (metadataBackedUp && Directory.Exists(metadataOld)) Directory.Move(metadataOld, MetadataPath);
                if (knowledgeBackedUp && Directory.Exists(old)) Directory.Move(old, KnowledgePath);
                throw;
            }
            DeleteDirectory(old); DeleteDirectory(metadataOld);
        }
        finally { DeleteDirectory(staging); DeleteDirectory(metadataNew); }
    }

    private async Task<string> DownloadExpectedShaAsync(string tag, CancellationToken ct)
    {
        var text = await _http.GetStringAsync(AssetUri(tag, ArchiveName + ".sha256"), ct);
        var match = Regex.Match(text, "^[0-9a-fA-F]{64}"); if (!match.Success) throw new InvalidDataException("Ungültige SHA256-Datei."); return match.Value.ToLowerInvariant();
    }
    internal async Task DownloadAsync(Uri uri, string destination, CancellationToken ct) { if (uri.Scheme != Uri.UriSchemeHttps) throw new InvalidOperationException("Nur HTTPS-Downloads sind erlaubt."); await using var input = await _http.GetStreamAsync(uri, ct); await using var output = new FileStream(destination, FileMode.Create, FileAccess.Write, FileShare.None); await input.CopyToAsync(output, ct); }
    private static Uri AssetUri(string tag, string asset) => new($"https://github.com/DerHausmeisterKrause/project_tool/releases/download/{Uri.EscapeDataString(tag)}/{asset}");
    private static string[] EntryParts(string name) => name.Replace('\\', '/').Split('/', StringSplitOptions.RemoveEmptyEntries);
    private static string SafePath(string root, string[] parts) { var fullRoot = Path.GetFullPath(root) + Path.DirectorySeparatorChar; var result = Path.GetFullPath(Path.Combine([root, .. parts])); if (!result.StartsWith(fullRoot, StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Unsicherer ZIP-Pfad."); return result; }
    private static async Task<string> CalculateShaAsync(string path, CancellationToken ct) { await using var stream = File.OpenRead(path); return Convert.ToHexString(await SHA256.HashDataAsync(stream, ct)).ToLowerInvariant(); }
    private async Task<StandardKnowledgeInstallation?> ReadInstalledAsync(CancellationToken ct)
    {
        var path = Path.Combine(MetadataPath, "installed.json");
        if (!File.Exists(path)) return null;
        try
        {
            await using var stream = File.OpenRead(path);
            var installation = ValidateInstallation(await JsonSerializer.DeserializeAsync<StandardKnowledgeInstallation>(stream, JsonOptions, ct));
            if (installation == null) LogInvalidMetadata();
            return installation;
        }
        catch (OperationCanceledException) { throw; }
        catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException)
        {
            LogInvalidMetadata(); return null;
        }
    }

    private async Task WriteInstalledAsync(StandardKnowledgeInstallation value, CancellationToken ct, string? directory = null)
    {
        directory ??= MetadataPath; Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "installed.json"); var temporary = path + ".tmp";
        try
        {
            await using (var stream = new FileStream(temporary, FileMode.Create, FileAccess.Write, FileShare.None))
                await JsonSerializer.SerializeAsync(stream, value, JsonOptions, ct);
            File.Move(temporary, path, true);
        }
        finally { if (File.Exists(temporary)) File.Delete(temporary); }
    }

    private void RefreshStatus(string? error = null)
    {
        StandardKnowledgeInstallation? installation = null;
        var path = Path.Combine(MetadataPath, "installed.json");
        if (File.Exists(path))
        {
            try
            {
                installation = ValidateInstallation(JsonSerializer.Deserialize<StandardKnowledgeInstallation>(File.ReadAllText(path), JsonOptions));
                if (installation == null) LogInvalidMetadata();
            }
            catch (Exception ex) when (ex is JsonException or IOException or UnauthorizedAccessException) { LogInvalidMetadata(); }
        }
        var count = 0;
        try { if (Directory.Exists(KnowledgePath)) count = Directory.EnumerateFiles(KnowledgePath, "*", SearchOption.AllDirectories).Count(p => SupportedExtensions.Contains(Path.GetExtension(p), StringComparer.OrdinalIgnoreCase)); }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException) { error ??= ex.Message; }
        SetStatus(error == null && installation != null ? StandardKnowledgeState.Installed : error == null ? StandardKnowledgeState.NotInstalled : StandardKnowledgeState.DownloadFailed, error, count, installation?.ReleaseTag);
    }

    private static StandardKnowledgeInstallation? ValidateInstallation(StandardKnowledgeInstallation? installation)
        => installation is { ReleaseTag.Length: > 0, ArchiveFile.Length: > 0, Sha256.Length: 64 }
           && installation.InstalledUtc != default && Regex.IsMatch(installation.Sha256, "^[0-9a-fA-F]{64}$") ? installation : null;

    private void LogInvalidMetadata() => _logger.Warning("[Standard Knowledge] installed metadata invalid; reinstall required");
    private void SetStatus(StandardKnowledgeState state, string? error = null, int? count = null, string? tag = null) { Status = new(state, count ?? Status.DocumentCount, tag ?? Status.ReleaseTag, error); StatusChanged?.Invoke(this, EventArgs.Empty); }
    private static void DeleteDirectory(string path) { if (Directory.Exists(path)) Directory.Delete(path, true); }
}
