using System.IO;
using Microsoft.Data.Sqlite;
using TaskTool.Models;
using TaskTool.Services;
using Xunit;

namespace TaskTool.Tests;

public sealed class AiKnowledgeTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "PlenaroKnowledgeTests", Guid.NewGuid().ToString("N"));
    private readonly LoggerService _logger = new(AppLogLevel.Error);

    [Fact] public void Knowledge_IsDisabledByDefault() => Assert.False(new AppSettings().AiKnowledgeEnabled);

    [Fact]
    public void Paths_AreBelowLocalPlenaroAiDirectory()
    {
        var index = new AiKnowledgeIndexService(_logger, localAppData: _root);
        Assert.Equal(Path.Combine(_root, "Plenaro", "AI", "knowledge"), index.KnowledgePath);
        Assert.Equal(Path.Combine(_root, "Plenaro", "AI", "knowledge-index.db"), index.IndexPath);
    }

    [Theory] [InlineData("a.txt", true)] [InlineData("a.MD", true)] [InlineData("a.PDF", true)] [InlineData("a.docx", false)]
    public void SupportedExtensions_AreFiltered(string name, bool expected) => Assert.Equal(expected, new AiKnowledgeDocumentExtractor().IsSupported(name));

    [Fact]
    public async Task TextAndMarkdown_AreExtracted()
    {
        Directory.CreateDirectory(_root); var extractor = new AiKnowledgeDocumentExtractor();
        foreach (var name in new[] { "wissen.txt", "wissen.md" }) { var path = Path.Combine(_root, name); await File.WriteAllTextAsync(path, "PLENARO-4711"); Assert.Contains("PLENARO-4711", (await extractor.ExtractAsync(path))[0].Text); }
    }

    [Fact]
    public async Task SmallTextPdf_IsExtracted()
    {
        Directory.CreateDirectory(_root); var path = Path.Combine(_root, "tiny.pdf");
        await File.WriteAllBytesAsync(path, CreateTinyPdf("PLENARO PDF TEST"));
        var pages = await new AiKnowledgeDocumentExtractor().ExtractAsync(path);
        Assert.Contains("PLENARO", string.Join(" ", pages.Select(x => x.Text)));
    }

    [Fact]
    public void Chunking_RespectsMaximumAndCreatesOverlap()
    {
        var chunks = AiKnowledgeChunker.Chunk(new[] { new ExtractedKnowledgeText(string.Join(' ', Enumerable.Repeat("abcdefghij", 400))) });
        Assert.True(chunks.Count > 1); Assert.All(chunks, chunk => Assert.InRange(chunk.Content.Length, 1, AiKnowledgeChunker.MaximumSize));
    }

    [Fact]
    public async Task Index_IsIncremental_UpdatesAndDeletesDocuments_AndFtsFindsContent()
    {
        var index = new AiKnowledgeIndexService(_logger, localAppData: _root); Directory.CreateDirectory(Path.Combine(index.KnowledgePath, "Windows"));
        var path = Path.Combine(index.KnowledgePath, "Windows", "GPO.txt"); await File.WriteAllTextAsync(path, "gruppenrichtlinie testserver PLENARO4711");
        var first = await index.IndexAsync(); var timestamp = await IndexedUtc(index.IndexPath); await index.IndexAsync(); Assert.Equal(timestamp, await IndexedUtc(index.IndexPath));
        var search = new AiKnowledgeSearchService(index.IndexPath, _logger); var result = await search.SearchAsync("Windows gruppenrichtlinie", 1);
        Assert.Single(result); Assert.Equal("Windows\\GPO.txt", result[0].RelativePath);
        await Task.Delay(20); await File.WriteAllTextAsync(path, "gruppenrichtlinie geändert PLENARO4712"); File.SetLastWriteTimeUtc(path, DateTime.UtcNow.AddSeconds(1)); await index.IndexAsync();
        Assert.Contains("PLENARO4712", (await search.SearchAsync("PLENARO4712"))[0].Content); Assert.NotEqual(timestamp, await IndexedUtc(index.IndexPath));
        File.Delete(path); Assert.Equal(0, (await index.IndexAsync()).DocumentCount);
    }

    [Fact]
    public void Context_HonorsCharacterBudgetAndTopNConstant()
    {
        var matches = Enumerable.Range(0, 10).Select(i => new AiKnowledgeMatch(new string('x', 1200), $"Windows\\{i}.txt", "Windows", $"{i}.txt", null, i)).Take(AiKnowledgeSearchService.DefaultTopN).ToArray();
        Assert.True(AiKnowledgeContextBuilder.Build(matches).Length <= AiKnowledgeContextBuilder.MaximumContextCharacters);
        Assert.Equal(4, AiKnowledgeSearchService.DefaultTopN);
    }

    [Fact]
    public void CombinedContext_RanksBothSourcesAndHonorsSharedBudget()
    {
        var matches = Enumerable.Range(0, 8).Select(index => new AiRetrievalMatch(
            index % 2 == 0 ? AiKnowledgeSourceType.Wiki : AiKnowledgeSourceType.LocalFiles,
            new string('x', 1400), $"Titel {index}", index % 2 == 0 ? "Internes Confluence" : $"Datei{index}.md", null, 100 - index)).ToArray();
        var context = AiCombinedContextBuilder.Prepare(matches);
        Assert.True(context.Text.Length <= AiKnowledgeContextBuilder.MaximumContextCharacters);
        Assert.True(context.IncludedMatches.Count <= AiKnowledgeContextBuilder.MaximumChunks);
        Assert.Contains(context.IncludedMatches, x => x.SourceType == AiKnowledgeSourceType.Wiki);
        Assert.Contains(context.IncludedMatches, x => x.SourceType == AiKnowledgeSourceType.LocalFiles);
    }

    [Theory]
    [InlineData("test")]
    [InlineData("Antworte nur mit test")]
    [InlineData("Hallo")]
    [InlineData("Danke")]
    public async Task Search_SkipsPromptsWithoutMeaningfulTerms(string question)
    {
        var search = await CreateSearchAsync(("Allgemein/Ports.md", "Server Port Test und weitere Informationen"));
        Assert.Empty(await search.SearchAsync(question));
    }

    [Fact]
    public void NormalizeTerms_RemovesConversationalFillersButKeepsTechnicalTerms()
    {
        var terms = AiKnowledgeSearchService.NormalizeTerms("Ich habe dieses Problem an meinem PC, was kann man dagegen mit gpupdate machen?");

        Assert.DoesNotContain(terms, term => new[] { "habe", "dieses", "meinem", "kann", "man", "dagegen", "machen" }.Contains(term, StringComparer.OrdinalIgnoreCase));
        Assert.Contains("pc", terms, StringComparer.OrdinalIgnoreCase);
        Assert.Contains("gpupdate", terms, StringComparer.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Search_AcceptsSpecificLinuxTokenWithoutFillingTopN()
    {
        var search = await CreateSearchAsync(
            ("Linux/Fehler/address_already_in_use.md", "EADDRINUSE Address already in use ss -tulpn lsof -i"),
            ("VMware/vCenter/host_disconnected.md", "VMware vCenter Server ist nicht erreichbar"));

        var result = await search.SearchAsync("Mein Linux Dienst meldet EADDRINUSE auf Port 3000.");

        var match = Assert.Single(result);
        Assert.Equal("Linux\\Fehler\\address_already_in_use.md", match.RelativePath);
    }

    [Fact]
    public async Task Search_ReturnsOnlyDocumentMatchingWindowsErrorCode()
    {
        var search = await CreateSearchAsync(
            ("Windows/Fehler/netzwerkpfad.md", "0x80070035 Netzwerkpfad nicht gefunden Windows SMB"),
            ("VMware/vCenter/host_disconnected.md", "Windows Server VMware vCenter Fehler"),
            ("Webserver/NGINX/start.md", "Windows Server nginx Problem"));

        var result = await search.SearchAsync("Was bedeutet 0x80070035?");

        var match = Assert.Single(result);
        Assert.Equal("Windows\\Fehler\\netzwerkpfad.md", match.RelativePath);
    }

    [Fact]
    public async Task Search_RejectsUnrelatedWindowsDocumentsForGeneralPerformanceQuestion()
    {
        var search = await CreateSearchAsync(
            ("Windows/Active_Directory/rpc_server_nicht_verfuegbar.md", "Windows Active Directory RPC Server nicht verfügbar"),
            ("Windows/Gruppenrichtlinien_GPO/gpupdate_fehler_sysvol.md", "Windows gpupdate Fehler beim Zugriff auf SYSVOL"));

        Assert.Empty(await search.SearchAsync("Mein Windows PC ist sehr langsam."));
    }

    [Fact]
    public async Task Search_ReturnsOnlyStrongMetadataMatchForSingleMeaningfulTerm()
    {
        var search = await CreateSearchAsync(
            ("Windows/Active_Directory/rpc_server_nicht_verfuegbar.md", "Windows Active Directory RPC Server nicht verfügbar"),
            ("Windows/Gruppenrichtlinien_GPO/gpupdate_fehler_sysvol.md", "Windows gpupdate Fehler beim Zugriff auf SYSVOL"),
            ("Windows/Performance/windows_pc_langsam.md", "Windows PC langsam CPU-Auslastung prüfen Arbeitsspeicher prüfen Datenträgerauslastung prüfen Task-Manager Autostart"));

        var match = Assert.Single(await search.SearchAsync("Mein Windows PC ist sehr langsam."));
        Assert.Equal("Windows\\Performance\\windows_pc_langsam.md", match.RelativePath);
    }

    [Fact]
    public async Task Search_StillFindsSpecificGpupdateSysvolFailure()
    {
        var search = await CreateSearchAsync(
            ("Windows/Active_Directory/rpc_server_nicht_verfuegbar.md", "Windows Active Directory RPC Server nicht verfügbar"),
            ("Windows/Gruppenrichtlinien_GPO/gpupdate_fehler_sysvol.md", "gpupdate schlägt mit SYSVOL Fehler fehl"));

        var match = Assert.Single(await search.SearchAsync("gpupdate schlägt mit SYSVOL Fehler fehl"));
        Assert.Equal("Windows\\Gruppenrichtlinien_GPO\\gpupdate_fehler_sysvol.md", match.RelativePath);
    }

    private async Task<AiKnowledgeSearchService> CreateSearchAsync(params (string Path, string Content)[] documents)
    {
        var index = new AiKnowledgeIndexService(_logger, localAppData: _root);
        foreach (var document in documents)
        {
            var path = Path.Combine(index.KnowledgePath, document.Path.Replace('/', Path.DirectorySeparatorChar));
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            await File.WriteAllTextAsync(path, document.Content);
        }
        await index.IndexAsync();
        return new AiKnowledgeSearchService(index.IndexPath, _logger);
    }

    private static async Task<string> IndexedUtc(string path) { await using var db = new SqliteConnection($"Data Source={path}"); await db.OpenAsync(); await using var cmd = db.CreateCommand(); cmd.CommandText = "SELECT indexed_utc FROM knowledge_documents LIMIT 1"; return (string)(await cmd.ExecuteScalarAsync())!; }
    private static byte[] CreateTinyPdf(string text)
    {
        var content = $"BT /F1 12 Tf 20 100 Td ({text}) Tj ET\n";
        var objects = new[] { "<< /Type /Catalog /Pages 2 0 R >>", "<< /Type /Pages /Kids [3 0 R] /Count 1 >>", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>", $"<< /Length {System.Text.Encoding.ASCII.GetByteCount(content)} >>\nstream\n{content}endstream", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>" };
        using var stream = new MemoryStream(); using var writer = new StreamWriter(stream, System.Text.Encoding.ASCII, leaveOpen: true) { NewLine = "\n" }; writer.WriteLine("%PDF-1.4"); writer.Flush();
        var offsets = new List<long> { 0 }; for (var i = 0; i < objects.Length; i++) { offsets.Add(stream.Position); writer.WriteLine($"{i + 1} 0 obj"); writer.WriteLine(objects[i]); writer.WriteLine("endobj"); writer.Flush(); }
        var xref = stream.Position; writer.WriteLine($"xref\n0 {objects.Length + 1}\n0000000000 65535 f "); foreach (var offset in offsets.Skip(1)) writer.WriteLine($"{offset:0000000000} 00000 n "); writer.WriteLine($"trailer\n<< /Size {objects.Length + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF"); writer.Flush(); return stream.ToArray();
    }
    public void Dispose() { try { if (Directory.Exists(_root)) Directory.Delete(_root, true); } catch { } }
}
