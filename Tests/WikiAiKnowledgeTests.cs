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

    private sealed class FakeAiChatService : IAiChatService
    {
        public bool IsEnabled => true;
        public bool CanChat => true;
        public string ProviderDescription => "Test";
        public string AvailabilityMessage => string.Empty;
        public event EventHandler? StateChanged;
        public Task<string> ChatAsync(IReadOnlyList<AiChatRequestMessage> messages, AiRequestOptions options, CancellationToken cancellationToken = default)
            => Task.FromResult("Antwort");
    }

    public void Dispose()
    {
        try { if (Directory.Exists(_root)) Directory.Delete(_root, true); } catch { }
    }
}
