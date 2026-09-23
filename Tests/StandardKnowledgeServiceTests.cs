using System.IO.Compression;
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

    private StandardKnowledgeService CreateService() => new(_root, () => "V3.3.1-dev-2", () => Task.CompletedTask, _logger);
    private string CreateArchive(params (string Name, string Content)[] entries)
    {
        Directory.CreateDirectory(_root); var path = Path.Combine(_root, Guid.NewGuid() + ".zip");
        using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
        foreach (var item in entries) { var entry = zip.CreateEntry(item.Name); using var writer = new StreamWriter(entry.Open()); writer.Write(item.Content); }
        return path;
    }
    public void Dispose() { try { if (Directory.Exists(_root)) Directory.Delete(_root, true); } catch { } }
}
