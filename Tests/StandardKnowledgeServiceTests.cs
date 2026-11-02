using System.IO.Compression;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using TaskTool.Models;
using TaskTool.Services;
using Xunit;

namespace TaskTool.Tests;

public sealed class StandardKnowledgeServiceTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "PlenaroStandardKnowledgeTests", Guid.NewGuid().ToString("N"));
    private readonly LoggerService _logger = new(AppLogLevel.Error);

    [Fact]
    public async Task Extraction_StripsKnowledgeRoot_SeparatesMetadata_AndPreservesUserFiles()
    {
        Directory.CreateDirectory(Path.Combine(_root, "knowledge"));
        var userFile = Path.Combine(_root, "knowledge", "meine_datei.md"); await File.WriteAllTextAsync(userFile, "unverändert");
        var archive = CreateArchive(("knowledge/Windows/test.md", "Testwissen"), ("MANIFEST.json", "{\"article_files_excluding_readme\":1}"), ("STRUCTURE.txt", "structure"));
        var service = CreateService();

        await service.ExtractSafelyAsync(archive, "V3.3.1-dev-2", new string('a', 64));

        Assert.True(File.Exists(Path.Combine(service.KnowledgePath, "Windows", "test.md")));
        Assert.False(Directory.Exists(Path.Combine(service.KnowledgePath, "knowledge")));
        Assert.True(File.Exists(Path.Combine(service.MetadataPath, "MANIFEST.json")));
        Assert.True(File.Exists(Path.Combine(service.MetadataPath, "STRUCTURE.txt")));
        Assert.False(File.Exists(Path.Combine(service.KnowledgePath, "MANIFEST.json")));
        Assert.Equal("unverändert", await File.ReadAllTextAsync(userFile));
    }

    [Fact]
    public async Task Extraction_RejectsZipSlip_AndKeepsExistingStandardKnowledge()
    {
        var service = CreateService(); Directory.CreateDirectory(service.KnowledgePath);
        var existing = Path.Combine(service.KnowledgePath, "existing.md"); await File.WriteAllTextAsync(existing, "keep");
        var archive = CreateArchive(("knowledge/test.md", "valid"), ("knowledge/../../evil.txt", "evil"));

        await Assert.ThrowsAsync<InvalidDataException>(() => service.ExtractSafelyAsync(archive, "V1", new string('b', 64)));

        Assert.Equal("keep", await File.ReadAllTextAsync(existing));
        Assert.False(File.Exists(Path.Combine(_root, "evil.txt")));
    }

    [Fact]
    public async Task CorruptInstalledMetadata_DoesNotCrashConstruction_AndCanBeRepaired()
    {
        Directory.CreateDirectory(Path.Combine(_root, "standard-knowledge"));
        Directory.CreateDirectory(Path.Combine(_root, "knowledge-default"));
        await File.WriteAllTextAsync(Path.Combine(_root, "standard-knowledge", "installed.json"), "{ kaputt");
        await File.WriteAllTextAsync(Path.Combine(_root, "knowledge-default", "old.md"), "old");
        var archive = CreateArchiveBytes(("knowledge/new.md", "new"));
        var handler = new FakeReleaseHandler(archive, Sha(archive));

        var service = CreateService(handler);

        Assert.Equal(StandardKnowledgeState.NotInstalled, service.Status.State);
        Assert.Equal(1, service.Status.DocumentCount);
        Assert.True(await service.EnsureInstalledAsync());
        Assert.Equal(StandardKnowledgeState.Installed, service.Status.State);
        Assert.True(File.Exists(Path.Combine(service.KnowledgePath, "new.md")));
    }

    [Fact]
    public async Task SameSha_UpdatesReleaseTagWithoutDownloadingArchive()
    {
        var sha = new string('a', 64);
        await SeedInstallationAsync("V2.4.9-dev-5", sha, "existing");
        var handler = new FakeReleaseHandler([], sha);
        var service = CreateService(handler, "V2.4.9-dev-6");

        Assert.False(await service.EnsureInstalledAsync());

        Assert.Equal(0, handler.ArchiveRequests);
        Assert.Equal("V2.4.9-dev-6", ReadInstallation().ReleaseTag);
        Assert.Equal("existing", await File.ReadAllTextAsync(Path.Combine(service.KnowledgePath, "existing.md")));
    }

    [Fact]
    public async Task NewSha_DownloadsVerifiesInstallsAndPreservesUserKnowledge()
    {
        await SeedInstallationAsync("V1", new string('a', 64), "old");
        Directory.CreateDirectory(Path.Combine(_root, "knowledge"));
        var userPath = Path.Combine(_root, "knowledge", "user.md"); await File.WriteAllTextAsync(userPath, "user-owned");
        var archive = CreateArchiveBytes(("knowledge/new.md", "new"));
        var handler = new FakeReleaseHandler(archive, Sha(archive));
        var reindexes = 0;
        var service = CreateService(handler, reindex: () => { reindexes++; return Task.CompletedTask; });

        Assert.True(await service.EnsureInstalledAsync());

        Assert.Equal(1, handler.ArchiveRequests);
        Assert.Equal(1, reindexes);
        Assert.Equal("new", await File.ReadAllTextAsync(Path.Combine(service.KnowledgePath, "new.md")));
        Assert.Equal("user-owned", await File.ReadAllTextAsync(userPath));
    }

    [Fact]
    public async Task ShaMismatch_KeepsExistingKnowledgeAndMetadata()
    {
        var oldSha = new string('a', 64); await SeedInstallationAsync("V1", oldSha, "old");
        var archive = CreateArchiveBytes(("knowledge/new.md", "new"));
        var service = CreateService(new FakeReleaseHandler(archive, new string('b', 64)));

        Assert.False(await service.EnsureInstalledAsync());

        Assert.Equal(StandardKnowledgeState.Sha256Failed, service.Status.State);
        Assert.Equal("old", await File.ReadAllTextAsync(Path.Combine(service.KnowledgePath, "existing.md")));
        Assert.Equal(oldSha, ReadInstallation().Sha256);
    }

    [Fact]
    public async Task InvalidNewArchive_KeepsExistingKnowledgeAndMetadata()
    {
        var oldSha = new string('a', 64); await SeedInstallationAsync("V1", oldSha, "old");
        var invalidArchive = Encoding.UTF8.GetBytes("not a zip");
        var service = CreateService(new FakeReleaseHandler(invalidArchive, Sha(invalidArchive)));

        Assert.False(await service.EnsureInstalledAsync());

        Assert.Equal("old", await File.ReadAllTextAsync(Path.Combine(service.KnowledgePath, "existing.md")));
        Assert.Equal("V1", ReadInstallation().ReleaseTag);
    }

    [Fact]
    public async Task OfflineCheck_KeepsExistingStandardKnowledgeAvailable()
    {
        await SeedInstallationAsync("V1", new string('a', 64), "available offline");
        var service = CreateService(new FakeReleaseHandler(new HttpRequestException("offline")));

        Assert.False(await service.EnsureInstalledAsync());

        Assert.Equal(1, service.Status.DocumentCount);
        Assert.Equal("available offline", await File.ReadAllTextAsync(Path.Combine(service.KnowledgePath, "existing.md")));
    }

    private StandardKnowledgeService CreateService() => new(_root, () => "V3.3.1-dev-2", () => Task.CompletedTask, _logger);
    private StandardKnowledgeService CreateService(HttpMessageHandler handler, string tag = "V2", Func<Task>? reindex = null)
        => new(_root, () => tag, reindex ?? (() => Task.CompletedTask), _logger, new HttpClient(handler));

    private async Task SeedInstallationAsync(string tag, string sha, string content)
    {
        Directory.CreateDirectory(Path.Combine(_root, "knowledge-default"));
        Directory.CreateDirectory(Path.Combine(_root, "standard-knowledge"));
        await File.WriteAllTextAsync(Path.Combine(_root, "knowledge-default", "existing.md"), content);
        await File.WriteAllTextAsync(Path.Combine(_root, "standard-knowledge", "installed.json"),
            JsonSerializer.Serialize(new StandardKnowledgeInstallation(tag, StandardKnowledgeService.ArchiveName, sha, DateTime.UtcNow), new JsonSerializerOptions(JsonSerializerDefaults.Web)));
    }

    private StandardKnowledgeInstallation ReadInstallation() => JsonSerializer.Deserialize<StandardKnowledgeInstallation>(
        File.ReadAllText(Path.Combine(_root, "standard-knowledge", "installed.json")), new JsonSerializerOptions(JsonSerializerDefaults.Web))!;

    private byte[] CreateArchiveBytes(params (string Name, string Content)[] entries)
    {
        using var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, true))
            foreach (var item in entries) { var entry = zip.CreateEntry(item.Name); using var writer = new StreamWriter(entry.Open()); writer.Write(item.Content); }
        return stream.ToArray();
    }

    private static string Sha(byte[] value) => Convert.ToHexString(SHA256.HashData(value)).ToLowerInvariant();
    private string CreateArchive(params (string Name, string Content)[] entries)
    {
        Directory.CreateDirectory(_root); var path = Path.Combine(_root, Guid.NewGuid() + ".zip");
        using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
        foreach (var item in entries) { var entry = zip.CreateEntry(item.Name); using var writer = new StreamWriter(entry.Open()); writer.Write(item.Content); }
        return path;
    }

    private sealed class FakeReleaseHandler : HttpMessageHandler
    {
        private readonly byte[] _archive;
        private readonly string _sha;
        private readonly Exception? _failure;
        public int ArchiveRequests { get; private set; }
        public FakeReleaseHandler(byte[] archive, string sha) { _archive = archive; _sha = sha; }
        public FakeReleaseHandler(Exception failure) { _archive = []; _sha = string.Empty; _failure = failure; }
        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            if (_failure != null) return Task.FromException<HttpResponseMessage>(_failure);
            if (request.RequestUri!.AbsolutePath.EndsWith(".sha256", StringComparison.Ordinal))
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent($"{_sha}  {StandardKnowledgeService.ArchiveName}\n") });
            ArchiveRequests++;
            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(_archive) });
        }
    }
    public void Dispose() { try { if (Directory.Exists(_root)) Directory.Delete(_root, true); } catch { } }
}
