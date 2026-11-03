using System.Net;
using System.Text;
using TaskTool.Models;
using TaskTool.Services;
using Xunit;

namespace TaskTool.Tests;

public sealed class TicketSystemServiceHttpTests
{
    [Fact]
    public async Task PostTicketSearchRecoversSessionExactlyOnceAndRewritesBody()
    {
        var handler = new RecordingHandler((request, occurrence) => request.RequestUri!.AbsolutePath switch
        {
            "/Session" => Json(HttpStatusCode.OK, occurrence == 1 ? "{\"SessionID\":\"old\"}" : "{\"SessionID\":\"new\"}"),
            "/TicketSearch" when occurrence == 1 => Json(HttpStatusCode.Unauthorized, Expired),
            "/TicketSearch" => Json(HttpStatusCode.OK, "{\"TicketID\":[42]}"),
            _ => throw new InvalidOperationException()
        });
        using var service = new TicketSystemService(handler, "https://znuny.test/Session");

        await service.SendForHttpRegressionTestAsync(HttpMethod.Post, "https://znuny.test/Session", "SessionCreate", "{}");
        await service.SendForHttpRegressionTestAsync(HttpMethod.Post, "https://znuny.test/TicketSearch", "TicketSearchOwner", "{\"SessionID\":\"old\"}");

        Assert.Equal(2, handler.Count("/Session"));
        Assert.Equal(2, handler.Count("/TicketSearch"));
        Assert.Contains("\"SessionID\":\"new\"", handler.Requests.Where(r => r.Path == "/TicketSearch").Last().Body);
    }

    [Fact]
    public async Task GetTicketGetRecoversSessionExactlyOnceAndRewritesQuery()
    {
        var handler = new RecordingHandler((request, occurrence) => request.RequestUri!.AbsolutePath switch
        {
            "/Session" => Json(HttpStatusCode.OK, "{\"SessionID\":\"new\"}"),
            "/Ticket/42" when occurrence == 1 => Json(HttpStatusCode.Unauthorized, Expired),
            "/Ticket/42" => Json(HttpStatusCode.OK, "{\"Ticket\":{\"TicketID\":42}}"),
            _ => throw new InvalidOperationException()
        });
        using var service = new TicketSystemService(handler, "https://znuny.test/Session");

        await service.SendForHttpRegressionTestAsync(HttpMethod.Get, "https://znuny.test/Ticket/42?SessionID=old", "TicketGetDetails");

        Assert.Equal(1, handler.Count("/Session"));
        Assert.Equal(2, handler.Count("/Ticket/42"));
        Assert.Contains("SessionID=new", handler.Requests.Where(r => r.Path == "/Ticket/42").Last().Uri);
    }

    [Theory]
    [InlineData("TicketCreate")]
    [InlineData("TicketUpdateReply")]
    [InlineData("TicketUpdateTimeBooking")]
    public async Task WritesAreNeverRetried(string stage)
    {
        var handler = new RecordingHandler((_, _) => Json(HttpStatusCode.Unauthorized, Expired));
        using var service = new TicketSystemService(handler, "https://znuny.test/Session");

        await Assert.ThrowsAsync<TaskTool.Models.ZnunyApiException>(() =>
            service.SendForHttpRegressionTestAsync(HttpMethod.Post, $"https://znuny.test/{stage}", stage, "{\"SessionID\":\"old\"}"));

        Assert.Single(handler.Requests);
        Assert.Equal(0, handler.Count("/Session"));
    }

    [Fact]
    public async Task AutomaticReadRecoveryIsNotStoppedByHistoricalSixtyRequestBudget()
    {
        var handler = new RecordingHandler((request, occurrence) => request.RequestUri!.AbsolutePath switch
        {
            "/Session" => Json(HttpStatusCode.OK, "{\"SessionID\":\"new\"}"),
            "/Ticket/42" when occurrence == 1 => Json(HttpStatusCode.Unauthorized, Expired),
            "/Ticket/42" => Json(HttpStatusCode.OK, "{}"),
            _ => throw new InvalidOperationException()
        });
        using var service = new TicketSystemService(handler, "https://znuny.test/Session");
        service.SetAutomaticTrafficForHttpRegressionTest(60);

        var exception = await Record.ExceptionAsync(() => service.SendForHttpRegressionTestAsync(
            HttpMethod.Get, "https://znuny.test/Ticket/42?SessionID=old", "TicketGetDetails"));

        Assert.Equal(2, handler.Count("/Ticket/42"));
        Assert.Equal(1, handler.Count("/Session"));
        Assert.Null(exception);
    }

    [Fact]
    public void TrafficHistoryIsBoundedAndSummaryIsRateLimited()
    {
        using var service = new TicketSystemService(new RecordingHandler((_, _) => Json(HttpStatusCode.OK, "{}")), "https://znuny.test/Session");
        service.EnableInfoLoggingForRegressionTest();
        var now = new DateTime(2026, 1, 1, 12, 0, 0, DateTimeKind.Utc);
        service.UtcNow = () => now;

        for (var index = 0; index < 20_000; index++)
            service.RecordGlobalTraffic(ZnunyRequestReason.ManualTicketRefresh, "TicketSearchOwner");

        Assert.Equal(TicketSystemService.MaximumTrafficEvents, service.TrafficEventCount);
        Assert.Equal(1, service.TrafficSummaryCount);

        now = now.AddMinutes(16);
        service.RecordGlobalTraffic(ZnunyRequestReason.ManualTicketRefresh, "TicketSearchOwner");
        Assert.Equal(1, service.TrafficEventCount);
        Assert.Equal(2, service.TrafficSummaryCount);
    }

    [Fact]
    public void DisabledInfoLoggingDoesNotRetainOrSummarizeTraffic()
    {
        using var service = new TicketSystemService(new RecordingHandler((_, _) => Json(HttpStatusCode.OK, "{}")), "https://znuny.test/Session");
        for (var index = 0; index < 20_000; index++)
            service.RecordGlobalTraffic(ZnunyRequestReason.FullTimerSync, "TicketGetDetails");

        Assert.Equal(0, service.TrafficEventCount);
        Assert.Equal(0, service.TrafficSummaryCount);
    }

    [Fact]
    public async Task CacheOnlyTicketContextNeverCreatesARequestWhenCacheIsIncomplete()
    {
        var path = Path.Combine(Path.GetTempPath(), $"plenaro-context-{Guid.NewGuid():N}.db");
        try
        {
            var database = new DatabaseService(new LoggerService(), path);
            database.Initialize();
            var cache = new TicketDetailCacheService(database);
            cache.Store(new TaskTool.Models.TicketBookingContext("42", "20260042", "4711", "A-1", [], [], "", [], null, "", "Cached", null, "", null, ""),
                "open", DateTime.UtcNow);
            var handler = new RecordingHandler((_, _) => throw new InvalidOperationException("No HTTP request expected."));
            using var service = new TicketSystemService(handler, "https://znuny.test/Session", cache);
            var task = new TaskTool.Models.TaskItem { Title = "Local", Tags = "ZnunyTicketID:42" };

            var context = await service.GetTicketBookingContextAsync(task);

            Assert.Equal("4711", context.CostCenterValue);
            Assert.Empty(handler.Requests);
        }
        finally { try { File.Delete(path); } catch { } }
    }

    private const string Expired = "{\"Error\":{\"ErrorCode\":\"SessionInvalid\",\"ErrorMessage\":\"Session expired\"}}";
    private static HttpResponseMessage Json(HttpStatusCode status, string json) => new(status)
    {
        Content = new StringContent(json, Encoding.UTF8, "application/json")
    };

    private sealed class RecordingHandler(Func<HttpRequestMessage, int, HttpResponseMessage> respond) : HttpMessageHandler
    {
        public List<(string Path, string Uri, string Body)> Requests { get; } = [];
        public int Count(string path) => Requests.Count(request => request.Path == path);

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            var body = request.Content == null ? string.Empty : await request.Content.ReadAsStringAsync(cancellationToken);
            var path = request.RequestUri!.AbsolutePath;
            Requests.Add((path, request.RequestUri.ToString(), body));
            return respond(request, Count(path));
        }
    }
}
