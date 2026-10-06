using TaskTool.Models;
using TaskTool.Services;
using TaskTool.ViewModels;
using Xunit;

namespace TaskTool.Tests;

public sealed class WikiAiKnowledgeTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), "PlenaroWikiAiTests", Guid.NewGuid().ToString("N"));
    private readonly LoggerService _logger = new(AppLogLevel.Error);

    [Fact]
    public async Task Search_RejectsUnrelatedWikiPagesForGeneralWindowsQuestion()
    {
        using var fixture = CreateFixture(
            ("rpc", "RPC Server nicht verfügbar", "Windows Active Directory RPC Server nicht verfügbar"),
            ("gpo", "gpupdate SYSVOL Fehler", "Windows Gruppenrichtlinie gpupdate Fehler beim Zugriff auf SYSVOL"));
        await fixture.Service.SyncAsync(fixture.Source, true);

        Assert.Empty(await fixture.Service.SearchAsync("Mein Windows PC ist sehr langsam."));
    }

    [Fact]
    public async Task Search_ReturnsOnlyMatchingWikiPerformancePage()
    {
        using var fixture = CreateFixture(
            ("rpc", "RPC Server nicht verfügbar", "Windows Active Directory RPC Server nicht verfügbar"),
            ("gpo", "gpupdate SYSVOL Fehler", "Windows Gruppenrichtlinie gpupdate Fehler beim Zugriff auf SYSVOL"),
            ("performance", "Windows PC langsam", "Windows PC langsam CPU-Auslastung prüfen Arbeitsspeicher prüfen Datenträgerauslastung prüfen Task-Manager Autostart"));
        await fixture.Service.SyncAsync(fixture.Source, true);

        var match = Assert.Single(await fixture.Service.SearchAsync("Mein Windows PC ist sehr langsam."));
        Assert.Equal("Windows PC langsam", match.Title);
    }

    [Fact]
    public async Task UseWiki_IsInactiveWithoutSourceAndRestoresPersistedPreferenceWhenSourceAppears()
    {
        Directory.CreateDirectory(_root);
        var settings = new SettingsService(_logger, Path.Combine(_root, "settings.json"));
        settings.Current.AiChatUseWiki = true;
        settings.Save();
        var provider = new FakeWikiKnowledgeProvider(("page", "Performance", "Windows langsam Performance"));
        using var wiki = new WikiAiKnowledgeService(settings, _logger, new[] { ("ConfluenceDataCenter", (IWikiKnowledgeProvider)provider) }, _root);
        var viewModel = new AiChatViewModel(new FakeAiChatService(), () => { }, settings: settings, wikiKnowledge: wiki);

        Assert.False(viewModel.UseWikiAvailable);
        Assert.False(viewModel.UseWiki);
        Assert.True(settings.Current.AiChatUseWiki);

        var source = ValidSource();
        settings.Current.WikiSources.Add(source);
        await wiki.SyncAsync(source, true);

        Assert.True(viewModel.UseWikiAvailable);
        Assert.True(viewModel.UseWiki);
    }

    [Fact]
    public async Task RestrictedScope_RejectsProviderPagesOutsideConfiguredSpaces()
    {
        Directory.CreateDirectory(_root); var settings = new SettingsService(_logger, Path.Combine(_root, "settings.json"));
        var source = ValidSource(); source.SearchAllSpaces = false; source.SpaceKeys = ["IT", "NETZ"]; settings.Current.WikiSources.Add(source);
        var provider = new ScopedProvider(("it", "IT page", "IT", "SCOPE_IDENTIFIER4711"), ("hr", "HR page", "HR", "SCOPE_IDENTIFIER4711"));
        using var service = new WikiAiKnowledgeService(settings, _logger, [("ConfluenceDataCenter", (IWikiKnowledgeProvider)provider)], _root);

        await service.SyncAsync(source, true);
        var result = await service.SearchAsync("SCOPE_IDENTIFIER4711");

        Assert.Single(result); Assert.Equal("IT", result[0].SpaceKey);
        Assert.Equal(1, service.GetStatus(source.Id).PageCount);
    }

    [Fact]
    public async Task ChangedScopeOrDisabledSource_CannotSearchStaleIndex()
    {
        using var fixture = CreateFixture(("page", "Scope page", "SCOPE_CHANGE4711"));
        await fixture.Service.SyncAsync(fixture.Source, true);
        Assert.Single(await fixture.Service.SearchAsync("SCOPE_CHANGE4711"));

        fixture.Source.SearchAllSpaces = false; fixture.Source.SpaceKeys = ["NETZ"];
        Assert.Empty(await fixture.Service.SearchAsync("SCOPE_CHANGE4711"));
        fixture.Source.Enabled = false;
        Assert.Empty(await fixture.Service.SearchAsync("SCOPE_CHANGE4711"));
    }

    [Fact]
    public void LegacySpaceKey_IsUsedWhenSpaceKeysAreEmpty()
    {
        var source = ValidSource(); source.SearchAllSpaces = false; source.SpaceKey = "LEGACY";
        Assert.Equal(["LEGACY"], WikiScopePolicy.GetSpaceKeys(source));
        Assert.True(WikiScopePolicy.AllowsSpace(source, "legacy"));
    }

    [Fact]
    public void StructuredParser_PreservesSectionsTablesListsAndCode()
    {
        const string markup = "<h2>Fehlerbehebung</h2><p>NGINX 502 Bad Gateway</p><table><tr><th>Port</th><th>Dienst</th></tr><tr><td>443</td><td>HTTPS</td></tr><tr><td>22</td><td>SSH</td></tr></table><ul><li>Java 21</li><li>8 GB RAM</li></ul><pre>listen 443 ssl;</pre>";
        var page = new WikiKnowledgePage("s", "p", "NGINX", "https://wiki/p", "IT", "1", null);
        var parsed = new ConfluenceKnowledgeParser().Parse(page, new("p", "NGINX", "", "1", null, markup));

        Assert.All(parsed.Blocks, block => Assert.Equal("Fehlerbehebung", block.SectionTitle));
        var table = Assert.Single(parsed.Blocks, block => block.ContentKind == WikiKnowledgeContentKind.Table);
        Assert.Contains("| Port | Dienst |", table.Content); Assert.Contains("| 443 | HTTPS |", table.Content); Assert.Contains("| 22 | SSH |", table.Content);
        Assert.Contains(parsed.Blocks, block => block.ContentKind == WikiKnowledgeContentKind.List && block.Content.Contains("- Java 21"));
        Assert.Contains(parsed.Blocks, block => block.ContentKind == WikiKnowledgeContentKind.Code && block.Content.Contains("listen 443 ssl;"));
    }

    [Fact]
    public void DrawIoExtractor_PreservesNodesAndRelations()
    {
        const string xml = "<mxGraphModel><root><mxCell id='1' vertex='1' value='Client'/><mxCell id='2' vertex='1' value='Reverse Proxy'/><mxCell id='3' vertex='1' value='Database'/><mxCell id='4' edge='1' source='1' target='2' value='HTTPS'/><mxCell id='5' edge='1' source='2' target='3'/></root></mxGraphModel>";
        var text = new WikiDrawIoExtractor().Extract(xml);
        Assert.Contains("Client", text); Assert.Contains("Reverse Proxy", text); Assert.Contains("Database", text);
        Assert.Contains("Client -> Reverse Proxy [HTTPS]", text); Assert.Contains("Reverse Proxy -> Database", text);
    }

    private Fixture CreateFixture(params (string Id, string Title, string Content)[] pages)
    {
        Directory.CreateDirectory(_root);
        var settings = new SettingsService(_logger, Path.Combine(_root, "settings.json"));
        var source = ValidSource();
        settings.Current.WikiSources.Add(source);
        var provider = new FakeWikiKnowledgeProvider(pages);
        var service = new WikiAiKnowledgeService(settings, _logger, new[] { ("ConfluenceDataCenter", (IWikiKnowledgeProvider)provider) }, _root);
        return new(service, source);
    }

    private static WikiSourceSettings ValidSource() => new()
    {
        Id = "wiki-test",
        Name = "Test Wiki",
        ProviderType = "ConfluenceDataCenter",
        BaseUrl = "https://wiki.example.test",
        AuthMode = "BearerToken",
        SearchAllSpaces = true
    };

    private sealed record Fixture(WikiAiKnowledgeService Service, WikiSourceSettings Source) : IDisposable
    {
        public void Dispose() => Service.Dispose();
    }

    private sealed class FakeWikiKnowledgeProvider(params (string Id, string Title, string Content)[] pages) : IWikiKnowledgeProvider
    {
        public Task<WikiKnowledgePageBatch> GetPagesAsync(WikiSourceSettings source, int offset, int limit, CancellationToken token)
        {
            var result = pages.Skip(offset).Take(limit)
                .Select(page => new WikiKnowledgePage(source.Id, page.Id, page.Title, $"https://wiki.example.test/{page.Id}", "IT", "1", null))
                .ToArray();
            return Task.FromResult(new WikiKnowledgePageBatch(result, offset + result.Length < pages.Length));
        }

        public Task<WikiKnowledgePageContent> GetPageContentAsync(WikiSourceSettings source, string externalId, CancellationToken token)
        {
            var page = pages.Single(candidate => candidate.Id == externalId);
            return Task.FromResult(new WikiKnowledgePageContent(page.Id, page.Title, page.Content, "1", null));
        }
    }

    private sealed class ScopedProvider(params (string Id, string Title, string Space, string Content)[] pages) : IWikiKnowledgeProvider
    {
        public Task<WikiKnowledgePageBatch> GetPagesAsync(WikiSourceSettings source, int offset, int limit, CancellationToken token)
        {
            var batch = pages.Skip(offset).Take(limit).Select(page => new WikiKnowledgePage(source.Id, page.Id, page.Title, $"https://wiki/{page.Id}", page.Space, "1", null)).ToArray();
            return Task.FromResult(new WikiKnowledgePageBatch(batch, offset + batch.Length < pages.Length));
        }
        public Task<WikiKnowledgePageContent> GetPageContentAsync(WikiSourceSettings source, string externalId, CancellationToken token)
        {
            var page = pages.Single(x => x.Id == externalId); return Task.FromResult(new WikiKnowledgePageContent(page.Id, page.Title, page.Content, "1", null));
        }
    }

    private sealed class FakeAiChatService : IAiChatService
    {
        public bool IsEnabled => true;
        public bool CanChat => true;
        public string ProviderDescription => "Test";
        public string AvailabilityMessage => string.Empty;
        public event EventHandler? StateChanged { add { } remove { } }
        public Task<string> ChatAsync(IReadOnlyList<AiChatRequestMessage> messages, AiRequestOptions options, CancellationToken cancellationToken = default)
            => Task.FromResult("Antwort");
    }

    public void Dispose()
    {
        try { if (Directory.Exists(_root)) Directory.Delete(_root, true); } catch { }
    }
}
