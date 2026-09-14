using System.Net;
using System.Text;
using TaskTool.Models;
using TaskTool.Services;
using TaskTool.ViewModels;
using Xunit;

namespace TaskTool.Tests;

public sealed class ZnunyAgentTests
{
    [Fact]
    public void AgentListRouteHasExpectedDefault()
        => Assert.Equal("/Ticket/Agent/List", new AppSettings().TicketSystemAgentListRoute);

    [Theory]
    [InlineData("Marcel Asböck", "Marcel Asböck", "Marcel", "Asböck", "m.asboeck", 12)]
    [InlineData("Marcel Asböck", "", "Marcel", "Asböck", "m.asboeck", 12)]
    [InlineData("m.asboeck", "", "", "", "m.asboeck", 12)]
    [InlineData("Agent 12", "", "", "", "", 12)]
    public void DisplayNameUsesDocumentedFallbacks(string expected, string name, string first, string last, string login, int id)
        => Assert.Equal(expected, new ZnunyAgent { UserId = id, Name = name, FirstName = first, LastName = last, Login = login }.DisplayName);

    [Theory]
    [InlineData("{\"Agents\":[{\"UserID\":2,\"Name\":\"zeta\"},{\"UserID\":1,\"Name\":\"Alpha\"},{\"UserID\":2,\"Name\":\"Duplicate\"}]}")]
    [InlineData("{\"Data\":{\"Agents\":[{\"UserID\":2,\"Name\":\"zeta\"},{\"UserID\":1,\"Name\":\"Alpha\"},{\"UserID\":2,\"Name\":\"Duplicate\"}]}}")]
    public void ParserAcceptsSupportedShapesSortsAndDeduplicates(string json)
    {
        var agents = TicketSystemService.ParseAgentListResponse(json);
        Assert.Equal([1, 2], agents.Select(agent => agent.UserId));
    }

    [Fact]
    public void MissingCurrentAgentsAreAddedSynthetically()
    {
        var context = Context(42, "Alter Benutzer", 43, "Historische Verantwortliche");
        var agents = TodayViewModel.AddCurrentAgents([], context);
        Assert.Equal("Alter Benutzer", Assert.Single(agents.Where(agent => agent.UserId == 42)).DisplayName);
        Assert.Equal("Historische Verantwortliche", Assert.Single(agents.Where(agent => agent.UserId == 43)).DisplayName);
    }

    [Fact]
    public async Task RepeatedAndParallelLoadsUseOneAgentRequestAndConfiguredRouteWithSession()
    {
        var handler = new AgentHandler();
        using var service = new TicketSystemService(handler, "https://znuny.test/Session");
        service.ConfigureAgentListForHttpRegressionTest("https://znuny.test/api", "/Custom/Agents");

        await Task.WhenAll(service.GetAgentsAsync(), service.GetAgentsAsync(), service.GetAgentsAsync());
        await service.GetAgentsAsync();

        Assert.Equal(1, handler.Count("/Custom/Agents"));
        Assert.Contains("SessionID=test-session", handler.Requests.Single(request => request.Path == "/Custom/Agents").Uri);
    }

    [Fact]
    public async Task ExpiredCacheReloadsAndFailedRefreshPreservesValidCache()
    {
        var now = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        var handler = new AgentHandler();
        using var service = new TicketSystemService(handler, "https://znuny.test/Session") { UtcNow = () => now };
        await service.GetAgentsAsync();
        now = now.AddHours(25);
        await service.GetAgentsAsync();
        Assert.Equal(2, handler.Count("/Ticket/Agent/List"));

        handler.FailAgents = true;
        var cached = await service.GetAgentsAsync(forceRefresh: true);
        Assert.Equal(2, cached.Count);
    }

    [Theory]
    [InlineData(12, null, "\"OwnerID\":12", "ResponsibleID")]
    [InlineData(null, 23, "\"ResponsibleID\":23", "OwnerID")]
    public async Task AssignmentUpdateSendsOnlyChangedField(int? ownerId, int? responsibleId, string included, string excluded)
    {
        var handler = new AgentHandler();
        using var service = new TicketSystemService(handler, "https://znuny.test/Session");
        var result = await service.UpdateTicketAssignmentAsync("42", ownerId, responsibleId);
        var update = handler.Requests.Single(request => request.Path == "/Ticket/42/Update");
        Assert.True(result.Success);
        Assert.Contains(included, update.Body);
        Assert.DoesNotContain(excluded, update.Body);
    }

    [Fact]
    public async Task AssignmentUpdateCombinesBothFieldsAndNoChangeSendsNothing()
    {
        var handler = new AgentHandler();
        using var service = new TicketSystemService(handler, "https://znuny.test/Session");
        Assert.True((await service.UpdateTicketAssignmentAsync("42", 12, 23)).Success);
        var update = handler.Requests.Single(request => request.Path == "/Ticket/42/Update");
        Assert.Contains("\"OwnerID\":12", update.Body);
        Assert.Contains("\"ResponsibleID\":23", update.Body);
        Assert.True((await service.UpdateTicketAssignmentAsync("42", null, null)).Success);
        Assert.Single(handler.Requests.Where(request => request.Path == "/Ticket/42/Update"));
        Assert.Empty(handler.Requests.Where(request => request.Path == "/Ticket/42"));
    }

    private static TicketBookingContext Context(int ownerId, string owner, int responsibleId, string responsible)
        => new("42", "2026000042", "", "", [], [], "", [], null, "", "Ticket",
            ownerId, owner, responsibleId, responsible);

    private sealed class AgentHandler : HttpMessageHandler
    {
        public List<(string Path, string Uri, string Body)> Requests { get; } = [];
        public bool FailAgents { get; set; }
        public int Count(string path) => Requests.Count(request => request.Path == path);

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            var path = request.RequestUri!.AbsolutePath;
            var body = request.Content == null ? string.Empty : await request.Content.ReadAsStringAsync(cancellationToken);
            Requests.Add((path, request.RequestUri.ToString(), body));
            var json = path == "/Session" ? "{\"SessionID\":\"test-session\"}"
                : path.EndsWith("/Update", StringComparison.Ordinal) ? "{\"TicketID\":42}"
                : "{\"Agents\":[{\"UserID\":2,\"Name\":\"Zeta\"},{\"UserID\":1,\"Name\":\"Alpha\"}]}";
            var status = FailAgents && path != "/Session" ? HttpStatusCode.InternalServerError : HttpStatusCode.OK;
            return new HttpResponseMessage(status) { Content = new StringContent(json, Encoding.UTF8, "application/json") };
        }
    }
}
