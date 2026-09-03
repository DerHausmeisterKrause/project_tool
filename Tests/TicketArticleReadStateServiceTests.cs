using TaskTool.Models;
using TaskTool.Services;
using Xunit;

namespace TaskTool.Tests;

public sealed class TicketArticleReadStateServiceTests : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), $"plenaro-read-{Guid.NewGuid():N}.db");
    private readonly DatabaseService _database;
    private readonly TicketDetailCacheService _cache;

    public TicketArticleReadStateServiceTests()
    {
        _database = new DatabaseService(new LoggerService(), _path);
        _database.Initialize();
        _cache = new TicketDetailCacheService(_database);
    }

    [Fact]
    public void CachedArticlesBecomeReadUpgradeBaseline()
    {
        _cache.Store(Context("4711", Article("10", 10), Article("11", 11)), "open", DateTime.UtcNow,
            TicketDetailFetchProfile.Full(20));
        var service = new TicketArticleReadStateService(_database);
        Assert.False(service.HasUnread("4711"));
        Assert.False(service.IsUnread("4711", "10"));
    }

    [Fact]
    public void FirstReliableFetchCreatesReadBaseline()
    {
        var service = new TicketArticleReadStateService(_database);
        service.ReconcileFetchedArticles("4711", [Article("100", 100), Article("101", 101), Article("102", 102)]);
        Assert.Equal(0, service.GetUnreadCount("4711"));
    }

    [Fact]
    public void NewArticlesAreIdempotentlyUnreadAndIndividuallyRead()
    {
        var service = new TicketArticleReadStateService(_database);
        service.ReconcileFetchedArticles("4711", [Article("100", 100), Article("101", 101), Article("102", 102)]);
        service.ReconcileFetchedArticles("4711", [Article("100", 100), Article("101", 101), Article("102", 102), Article("103", 103)]);
        service.ReconcileFetchedArticles("4711", [Article("100", 100), Article("101", 101), Article("102", 102), Article("103", 103)]);
        Assert.Equal("103", Assert.Single(service.GetUnreadArticleIds("4711")));

        service.ReconcileFetchedArticles("4711", [Article("103", 103), Article("104", 104)]);
        Assert.Equal(2, service.GetUnreadCount("4711"));
        service.MarkRead("4711", "103");
        Assert.True(service.HasUnread("4711"));
        service.MarkRead("4711", "104");
        Assert.False(service.HasUnread("4711"));
    }

    [Fact]
    public void ReadStateSurvivesServiceRestart()
    {
        var first = new TicketArticleReadStateService(_database);
        first.ReconcileFetchedArticles("4711", [Article("10", 10)]);
        first.ReconcileFetchedArticles("4711", [Article("11", 11)]);
        Assert.True(new TicketArticleReadStateService(_database).IsUnread("4711", "11"));
    }

    [Fact]
    public void OlderArticlesFromIncreasedLimitAreReadAndWatermarkNeverMovesBackwards()
    {
        var service = new TicketArticleReadStateService(_database);
        service.ReconcileFetchedArticles("4711", [Article("20", 20)]);
        service.ReconcileFetchedArticles("4711", [Article("10", 10), Article("20", 20)]);
        Assert.False(service.IsUnread("4711", "10"));
        service.ReconcileFetchedArticles("4711", [Article("21", 21)]);
        Assert.True(service.IsUnread("4711", "21"));
    }

    [Fact]
    public void MissingArticleIdIsIgnoredAndDropdownUsesLocalStateOnly()
    {
        var service = new TicketArticleReadStateService(_database);
        service.ReconcileFetchedArticles("4711", [Article("10", 10), Article("", 11)]);
        service.ReconcileFetchedArticles("4711", [Article("11", 11)]);
        var article = Article("11", 11);
        article.IsUnread = service.IsUnread("4711", "11");
        Assert.StartsWith("★ ", article.DropdownDisplayText);
        service.MarkRead("4711", "11"); article.IsUnread = false;
        Assert.False(article.DropdownDisplayText.StartsWith("★", StringComparison.Ordinal));
    }

    private static TicketArticleItem Article(string id, int minute) => new()
    {
        ArticleId = id,
        CreatedLocal = new DateTime(2026, 9, 3, 10, minute % 60, 0, DateTimeKind.Utc),
        Body = "message",
        DisplayText = id
    };

    private static TicketBookingContext Context(string id, params TicketArticleItem[] articles)
        => new(id, id, "", "", [], [], "", articles, articles.LastOrDefault(), "customer@example.test", "Ticket");

    public void Dispose() { try { File.Delete(_path); } catch { } }
}
