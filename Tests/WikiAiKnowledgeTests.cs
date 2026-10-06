using TaskTool.Models;
using TaskTool.Services;
using TaskTool.ViewModels;
using Microsoft.Data.Sqlite;
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
    public async Task Search_AcceptsTeleportPageButRejectsPagesWithOnlyGenericOverlap()
    {
        using var fixture = CreateFixture(
            ("teleport", "Teleport Reverse Proxy", "Teleport Proxy wird intern über Port 3080 verwendet."),
            ("printer", "Druckerserver", "Benutzer benötigen Zugriff auf den Server."),
            ("files", "Fileserver", "Interner Zugriff auf Dateien."));
        await fixture.Service.SyncAsync(fixture.Source, true);

        var match = Assert.Single(await fixture.Service.SearchAsync("Teleport Proxy interner Zugriff"));
        Assert.Equal("Teleport Reverse Proxy", match.Title);
    }

    [Fact]
    public async Task Search_RejectsBestCandidateWhenEveryPageIsWeak()
    {
        using var fixture = CreateFixture(
            ("linux", "Linux Server", "Server Fehler beim Start."),
            ("files", "Fileserver", "Ein Fehler wurde protokolliert."));
        await fixture.Service.SyncAsync(fixture.Source, true);

        Assert.Empty(await fixture.Service.SearchAsync("Wie behebe ich einen Fehler am Drucker?"));
    }

    [Fact]
    public async Task Search_ReturnsBothComplementaryStrongPagesWithoutFillingToThree()
    {
        using var fixture = CreateFixture(
            ("proxy", "Teleport Proxy", "Teleport Proxy interner Zugriff Port 3080."),
            ("cert", "Teleport Zertifikat", "Teleport Proxy interner Zugriff benötigt ein Zertifikat."),
            ("printer", "Druckerserver", "Benutzer benötigen Zugriff auf den Server."));
        await fixture.Service.SyncAsync(fixture.Source, true);

        var matches = await fixture.Service.SearchAsync("Teleport Proxy interner Zugriff");
        Assert.Equal(2, matches.Select(match => match.ExternalPageId).Distinct().Count());
        Assert.DoesNotContain(matches, match => match.Title == "Druckerserver");
    }

    [Fact]
    public async Task Search_PwsAnchorRejectsUnrelatedUserAndPasswordPages()
    {
        using var fixture = CreateFixture(
            ("pws", "Password Safe Benutzerverwaltung", "PWS Benutzer können entsperrt werden."),
            ("fav", "FAV Neue FAV-User anlegen", "Benutzer anlegen und Passwort setzen."),
            ("mail", "mail-relay", "User benötigt Zugriff."),
            ("isms", "ISMS Runde", "Passwort Richtlinien."));
        await fixture.Service.SyncAsync(fixture.Source, true);

        var match = Assert.Single(await fixture.Service.SearchAsync("wie entsperre ich einen user im Password Safe PWS"));
        Assert.Equal("pws", match.ExternalPageId);
    }

    [Fact]
    public async Task Search_PentahoAnchorRejectsGenericLicensePagesAndReturnsZeroWithoutEntityPage()
    {
        using var fixture = CreateFixture(
            ("pentaho", "Pentaho Lizenzierung", "Lizenzdatei auf dem Pentaho Server einspielen."),
            ("windows", "Windows Server Lizenz", "Lizenz am Server einspielen."),
            ("isms", "ISMS Runde", "Server Lizenzen und Software."));
        await fixture.Service.SyncAsync(fixture.Source, true);

        var match = Assert.Single(await fixture.Service.SearchAsync("Pentaho Lizenz am Server einspielen"));
        Assert.Equal("pentaho", match.ExternalPageId);

        using var noEntityFixture = CreateFixture(
            ("windows", "Windows Server Lizenz", "Lizenz am Server einspielen."),
            ("isms", "ISMS Runde", "Server Lizenzen und Software."));
        await noEntityFixture.Service.SyncAsync(noEntityFixture.Source, true);
        Assert.Empty(await noEntityFixture.Service.SearchAsync("Pentaho Lizenz am Server einspielen"));
    }

    [Fact]
    public async Task Search_GeneralLinuxQuestionStillUsesNormalTechnicalRetrieval()
    {
        using var fixture = CreateFixture(("linux", "Linux Speicher prüfen", "Freien Speicher unter Linux mit df prüfen."));
        await fixture.Service.SyncAsync(fixture.Source, true);

        Assert.Single(await fixture.Service.SearchAsync("Wie prüfe ich freien Speicher unter Linux?"));
    }

    [Fact]
    public async Task Search_UsesAncestorHierarchyForPasswordSecureChildAndRejectsFavChild()
    {
        Directory.CreateDirectory(_root); var settings = new SettingsService(_logger, Path.Combine(_root, "settings.json")); var source = ValidSource(); settings.Current.WikiSources.Add(source);
        var provider = new HierarchyProvider(
            new("password-child", "Administrator Kennwort und 2FA zurücksetzen", "Öffne die Benutzerverwaltung und wähle den betroffenen Benutzer.", [new("root", "Password Secure"), new("admin", "Berechtigung")]),
            new("fav-child", "Benutzer anlegen", "Benutzer Kennwort zurücksetzen.", [new("fav", "FAV")]),
            new("generic-child", "Benutzerverwaltung", "Benutzer Kennwort zurücksetzen.", [new("apps", "Anwendungen")]));
        using var service = new WikiAiKnowledgeService(settings, _logger, [("ConfluenceDataCenter", (IWikiKnowledgeProvider)provider)], _root);
        await service.SyncAsync(source, true);

        var match = Assert.Single(await service.SearchAsync("Wie setze ich in Password Secure das Administrator Kennwort zurück?"));

        Assert.Equal("password-child", match.ExternalPageId);
        Assert.Equal("Password Secure > Berechtigung", match.HierarchyPath);
        Assert.Contains("Pfad: Password Secure > Berechtigung", AiCombinedContextBuilder.Prepare([match]).Text);
    }

    [Fact]
    public async Task Search_UsesGrandparentAndDetectsHierarchyMoveWithoutBodyVersionChange()
    {
        Directory.CreateDirectory(_root); var settings = new SettingsService(_logger, Path.Combine(_root, "settings.json")); var source = ValidSource(); settings.Current.WikiSources.Add(source);
        var provider = new HierarchyProvider(new("child", "2FA zurücksetzen", "Wähle den betroffenen Benutzer.", [new("general", "Allgemein")]));
        using var service = new WikiAiKnowledgeService(settings, _logger, [("ConfluenceDataCenter", (IWikiKnowledgeProvider)provider)], _root);
        await service.SyncAsync(source, true);
        Assert.Empty(await service.SearchAsync("Password Secure 2FA zurücksetzen"));

        provider.Pages[0] = provider.Pages[0] with { Ancestors = [new("password", "Password Secure"), new("admin", "Administration")] };
        await service.SyncAsync(source, false);

        var match = Assert.Single(await service.SearchAsync("Password Secure 2FA zurücksetzen"));
        Assert.Equal("Password Secure > Administration", match.HierarchyPath);
    }

    [Fact]
    public async Task Search_AttachmentInheritsParentPageHierarchy()
    {
        Directory.CreateDirectory(_root); var settings = new SettingsService(_logger, Path.Combine(_root, "settings.json")); var source = ValidSource(); settings.Current.WikiSources.Add(source);
        var provider = new AttachmentProvider { Ancestors = [new("password", "Password Secure")] };
        var processor = new WikiAttachmentProcessor(new FakePdfExtractor(), new FakeOcr("Button Entsperren auswählen"), new WikiDrawIoExtractor());
        using var service = new WikiAiKnowledgeService(settings, _logger, [("ConfluenceDataCenter", (IWikiKnowledgeProvider)provider)], _root, processor);
        await service.SyncAsync(source, true);

        var match = Assert.Single((await service.SearchAsync("Password Secure entsperren")).Where(result => result.ContentKind == WikiKnowledgeContentKind.ImageOcr));
        Assert.Equal(WikiKnowledgeContentKind.ImageOcr, match.ContentKind);
        Assert.Equal("Password Secure", match.HierarchyPath);
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

    [Fact]
    public async Task ReferencedPdfImageAndDrawIoAttachments_AreIndexedAndUnchangedPageIsNotDownloadedAgain()
    {
        Directory.CreateDirectory(_root); var settings = new SettingsService(_logger, Path.Combine(_root, "settings.json")); var source = ValidSource(); settings.Current.WikiSources.Add(source);
        var provider = new AttachmentProvider(); var pdfExtractor=new FakePdfExtractor(); var processor = new WikiAttachmentProcessor(pdfExtractor, new FakeOcr(), new WikiDrawIoExtractor());
        using var service = new WikiAiKnowledgeService(settings, _logger, [("ConfluenceDataCenter", (IWikiKnowledgeProvider)provider)], _root, processor);

        await service.SyncAsync(source, true);

        var pdf = Assert.Single(await service.SearchAsync("PDFORIGINAL4711")); Assert.Equal(WikiKnowledgeContentKind.Pdf, pdf.ContentKind); Assert.Equal("manual.pdf", pdf.AttachmentName); Assert.Equal(2, pdf.PageNumber);
        Assert.Equal(WikiKnowledgeContentKind.ImageOcr, Assert.Single(await service.SearchAsync("PLENARO-OCR-4711")).ContentKind);
        Assert.Equal(WikiKnowledgeContentKind.DrawIo, Assert.Single(await service.SearchAsync("Reverse Proxy Database")).ContentKind);
        Assert.Equal(3, provider.DownloadCount);
        Assert.DoesNotContain(
            Directory.EnumerateFiles(_root, "*", SearchOption.AllDirectories),
            path => new[] { ".pdf", ".png", ".drawio" }.Contains(Path.GetExtension(path), StringComparer.OrdinalIgnoreCase));
        await service.SyncAsync(source, false);
        Assert.Equal(3, provider.DownloadCount);
        var status = service.GetStatus(source.Id); Assert.Equal(3, status.AttachmentCount); Assert.Equal(1, status.PdfCount); Assert.Equal(1, status.OcrSuccessCount); Assert.Equal(0, status.OcrFailureCount); Assert.Equal(1, status.DrawIoCount);

        provider.PdfVersion = "2"; provider.PdfText = "PDFUPDATED4712";
        await service.SyncAsync(source, false);
        Assert.Equal(4, provider.DownloadCount);
        Assert.Single(await service.SearchAsync("PDFUPDATED4712"));
        Assert.Empty(await service.SearchAsync("PDFORIGINAL4711"));
        Assert.Equal(2,pdfExtractor.CallCount);
        provider.PdfVersion="3"; await service.SyncAsync(source,false);
        Assert.Equal(5,provider.DownloadCount); Assert.Equal(2,pdfExtractor.CallCount);
        await using(var db=new SqliteConnection($"Data Source={service.IndexPath}")){await db.OpenAsync();await using var command=db.CreateCommand();command.CommandText="UPDATE wiki_ai_attachments SET extraction_version=0 WHERE attachment_id='pdf'";await command.ExecuteNonQueryAsync();}
        provider.PdfVersion="4";await service.SyncAsync(source,false);Assert.Equal(6,provider.DownloadCount);Assert.Equal(3,pdfExtractor.CallCount);
    }

    [Fact]
    public void WindowsOcrService_DoesNotCreatePersistentTempFiles()
    {
        _ = new WindowsWikiImageOcrService(_root);
        Assert.False(Directory.Exists(Path.Combine(_root,"Plenaro","AI","temp")));
    }

    [Fact]
    public async Task GenericXmlAttachment_IsNotParsedAsDrawIo()
    {
        var processor = new WikiAttachmentProcessor(new FakePdfExtractor(), new FakeOcr(), new WikiDrawIoExtractor());
        var page = new WikiKnowledgePage("source", "page", "Page", "https://wiki/page", "IT", "1", null);
        var metadata = new WikiKnowledgeAttachmentMetadata("xml", "page", "settings.xml", "application/xml", "1", null, "https://wiki/settings.xml", 32);
        await using var stream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes("<settings><value>test</value></settings>"));

        var result = await processor.ProcessAsync(page, metadata, stream, CancellationToken.None);

        Assert.Equal("unsupported", result.Status);
        Assert.Equal(WikiKnowledgeContentKind.AttachmentMetadata, Assert.Single(result.Blocks).ContentKind);
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

    private sealed class AttachmentProvider : IWikiKnowledgeProvider
    {
        public int DownloadCount { get; private set; }
        public string PdfVersion { get; set; } = "1";
        public string PdfText { get; set; } = "PDFORIGINAL4711";
        public IReadOnlyList<WikiKnowledgeAncestor> Ancestors { get; set; } = [];
        public Task<WikiKnowledgePageBatch> GetPagesAsync(WikiSourceSettings source, int offset, int limit, CancellationToken token) => Task.FromResult(new WikiKnowledgePageBatch(offset == 0 ? [new(source.Id,"p","Serverbetrieb","https://wiki/p","IT","1",null,Ancestors)] : [], false));
        public Task<WikiKnowledgePageContent> GetPageContentAsync(WikiSourceSettings source,string externalId,CancellationToken token) => Task.FromResult(new WikiKnowledgePageContent("p","Serverbetrieb","Seitentext","1",null,"<p>Seitentext</p><ac:image><ri:attachment ri:filename='screen.png'/></ac:image><a><ri:attachment ri:filename='manual.pdf'/></a><ac:structured-macro ac:name='drawio'><ac:parameter>network.drawio</ac:parameter></ac:structured-macro>"));
        public Task<IReadOnlyList<WikiKnowledgeAttachmentMetadata>> GetAttachmentsAsync(WikiSourceSettings source,string pageId,CancellationToken token) => Task.FromResult<IReadOnlyList<WikiKnowledgeAttachmentMetadata>>([
            new("pdf","p","manual.pdf","application/pdf",PdfVersion,null,"https://wiki/manual.pdf",10), new("img","p","screen.png","image/png","1",null,"https://wiki/screen.png",10), new("draw","p","network.drawio","application/vnd.jgraph.mxfile","1",null,"https://wiki/network.drawio",100)]);
        public Task<Stream> DownloadAttachmentAsync(WikiSourceSettings source,WikiKnowledgeAttachmentMetadata attachment,CancellationToken token) { DownloadCount++; var value=attachment.FileName.EndsWith(".drawio")?"<mxGraphModel><root><mxCell id='1' vertex='1' value='Reverse Proxy'/><mxCell id='2' vertex='1' value='Database'/><mxCell id='3' edge='1' source='1' target='2'/></root></mxGraphModel>":attachment.FileName.EndsWith(".pdf")?PdfText:"bytes";return Task.FromResult<Stream>(new MemoryStream(System.Text.Encoding.UTF8.GetBytes(value))); }
    }
    private sealed class FakePdfExtractor : IWikiPdfExtractor { public int CallCount{get;private set;} public Task<IReadOnlyList<ExtractedKnowledgeText>> ExtractAsync(byte[] content,CancellationToken token){CallCount++;return Task.FromResult<IReadOnlyList<ExtractedKnowledgeText>>([new("Allgemein",1),new(System.Text.Encoding.UTF8.GetString(content),2)]);} }
    private sealed class FakeOcr(string text = "PLENARO-OCR-4711") : IWikiImageOcrService { public bool IsAvailable=>true; public Task<string> RecognizeAsync(byte[] content,string mediaType,CancellationToken token)=>Task.FromResult(text); }

    private sealed record HierarchyPage(string Id, string Title, string Content, IReadOnlyList<WikiKnowledgeAncestor> Ancestors);
    private sealed class HierarchyProvider(params HierarchyPage[] pages) : IWikiKnowledgeProvider
    {
        public HierarchyPage[] Pages { get; } = pages;
        public Task<WikiKnowledgePageBatch> GetPagesAsync(WikiSourceSettings source, int offset, int limit, CancellationToken token)
        {
            var result = Pages.Skip(offset).Take(limit).Select(page => new WikiKnowledgePage(source.Id, page.Id, page.Title, $"https://wiki/{page.Id}", "IT", "1", null, page.Ancestors)).ToArray();
            return Task.FromResult(new WikiKnowledgePageBatch(result, offset + result.Length < Pages.Length));
        }
        public Task<WikiKnowledgePageContent> GetPageContentAsync(WikiSourceSettings source, string externalId, CancellationToken token)
        {
            var page = Pages.Single(candidate => candidate.Id == externalId);
            return Task.FromResult(new WikiKnowledgePageContent(page.Id, page.Title, page.Content, "1", null));
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
