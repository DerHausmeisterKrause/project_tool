using System.Reflection;
using TaskTool.Models;
using TaskTool.Services;
using TaskTool.ViewModels;
using System.Net;
using System.Text;

namespace TaskTool.Tests;

public sealed class AgentAvailabilityTests
{
    [Theory]
    [InlineData(.20, null, AgentAvailabilityLevel.Green)]
    [InlineData(.34, null, AgentAvailabilityLevel.Green)]
    [InlineData(.35, null, AgentAvailabilityLevel.Orange)]
    [InlineData(.64, null, AgentAvailabilityLevel.Orange)]
    [InlineData(.65, null, AgentAvailabilityLevel.Red)]
    [InlineData(.90, 0, AgentAvailabilityLevel.Red)]
    [InlineData(.30, 10, AgentAvailabilityLevel.Orange)]
    [InlineData(.55, 10, AgentAvailabilityLevel.Orange)]
    [InlineData(.55, 15, AgentAvailabilityLevel.Red)]
    [InlineData(.10, 20, AgentAvailabilityLevel.Green)]
    [InlineData(null, 20, AgentAvailabilityLevel.Unknown)]
    public void Classify_RequiresOutlookAndReturnsOnlyCoarseLevel(double? calendarLoad, int? tickets, AgentAvailabilityLevel expected)
        => Assert.Equal(expected, AgentAvailabilityService.Classify(calendarLoad, tickets));

    [Theory]
    [InlineData(.30)]
    [InlineData(.55)]
    public void Classify_IncreasingTicketCountNeverImprovesAvailability(double calendarLoad)
    {
        var levels = new[] { 0, 10, 20 }
            .Select(tickets => AgentAvailabilityService.Classify(calendarLoad, tickets))
            .ToArray();

        Assert.True(levels[0] <= levels[1]);
        Assert.True(levels[1] <= levels[2]);
    }

    [Fact]
    public void FreeBusyParser_UsesOnlyFutureWeekdayPlanningSlots()
    {
        var monday = new DateTime(2026, 1, 5, 0, 0, 0, DateTimeKind.Local);
        var slots = Enumerable.Repeat('0', 14 * 48).ToArray();
        // Monday 08:00 tentative, 08:30 busy, 09:00 OOF; weekend busy must not count.
        slots[16] = '1'; slots[17] = '2'; slots[18] = '3'; slots[5 * 48 + 16] = '2';
        Assert.True(AgentAvailabilityService.TryCalculateCalendarLoad(
            new OutlookFreeBusySnapshot(monday, 30, new string(slots)), monday.AddHours(7), out var load));
        Assert.Equal(2.5 / (10 * 18), load, 6);
    }

    [Fact]
    public async Task CacheAndInflight_DeduplicateAndExpireWithoutRetainingInputs()
    {
        var now = new DateTime(2026, 1, 5, 8, 0, 0, DateTimeKind.Utc);
        var calls = 0;
        var gate = new TaskCompletionSource<OutlookFreeBusySnapshot?>(TaskCreationOptions.RunContinuationsAsynchronously);
        var service = new AgentAvailabilityService(
            _ => { Interlocked.Increment(ref calls); return gate.Task; },
            _ => Task.FromResult<int?>(null),
            () => now,
            () => now.ToLocalTime());
        var agent = new ZnunyAgent { UserId = 7, Name = "Not retained in cache" };

        var first = service.GetAvailabilityAsync(agent);
        var second = service.GetAvailabilityAsync(agent);
        gate.SetResult(AllFree(now.ToLocalTime()));
        Assert.Equal(AgentAvailabilityLevel.Green, (await first).Level);
        Assert.Equal((await first).Level, (await second).Level);
        Assert.Equal(1, calls);
        Assert.Equal(AgentAvailabilityLevel.Green, (await service.GetAvailabilityAsync(agent)).Level);
        Assert.Equal(1, calls);

        now = now.AddMinutes(11);
        Assert.Equal(AgentAvailabilityLevel.Green, (await service.GetAvailabilityAsync(agent)).Level);
        Assert.Equal(2, calls);
    }

    [Fact]
    public void Snapshot_ExposesNoDetailedOrPerformanceData()
    {
        var forbidden = new[] { "Percent", "Score", "Ticket", "Calendar", "BusyHours", "Hours" };
        var names = typeof(AgentAvailabilitySnapshot).GetProperties(BindingFlags.Instance | BindingFlags.Public).Select(p => p.Name);
        Assert.DoesNotContain(names, name => forbidden.Any(term => name.Contains(term, StringComparison.OrdinalIgnoreCase)));
    }

    [Theory]
    [InlineData("Email")]
    [InlineData("EmailAddress")]
    [InlineData("UserEmail")]
    [InlineData("Mail")]
    public void AgentParser_AcceptsOptionalEmailFields(string field)
    {
        var agents = TicketSystemService.ParseAgentListResponse($$"""{"Agents":[{"UserID":3,"Name":"Agent","{{field}}":"agent@example.test"}]}""");
        Assert.Equal("agent@example.test", Assert.Single(agents).Email);
    }

    [Fact]
    public void TaskDetail_ShowsOnlyCoarseAvailabilityBindings()
    {
        var xaml = File.ReadAllText(Path.Combine(RepoRoot(), "Views", "TaskDetailPanel.xaml"));
        Assert.Contains("Verfügbarkeit nächste 14 Tage", xaml, StringComparison.Ordinal);
        Assert.DoesNotContain("AvailabilityPercent", xaml, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("TicketCount", xaml, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("CalendarLoad", xaml, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void StaleSelection_CannotReplaceCurrentAgent()
    {
        var requested = new ZnunyAgent { UserId = 1 };
        Assert.False(TodayViewModel.IsCurrentAvailabilitySelection(new ZnunyAgent { UserId = 2 }, requested));
        Assert.True(TodayViewModel.IsCurrentAvailabilitySelection(new ZnunyAgent { UserId = 1 }, requested));
    }

    [Fact]
    public async Task OpenTicketCount_UsesOnlySearchAndDeduplicatesOwnerAndResponsibleIds()
    {
        var handler = new AvailabilityTicketHandler();
        using var service = new TicketSystemService(handler, "https://znuny.test/Session");
        var count = await service.TryGetAgentOpenTicketCountAsync(42);

        Assert.Equal(4, count);
        Assert.Equal(4, handler.Requests.Count(path => path == "/Ticket"));
        Assert.DoesNotContain(handler.Requests, path => path.StartsWith("/Ticket/", StringComparison.Ordinal));
    }

    private static OutlookFreeBusySnapshot AllFree(DateTime start)
        => new(start.Date, 30, new string('0', 14 * 48));

    private static string RepoRoot()
        => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));

    private sealed class AvailabilityTicketHandler : HttpMessageHandler
    {
        public List<string> Requests { get; } = [];

        protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            var path = request.RequestUri!.AbsolutePath;
            Requests.Add(path);
            if (path.EndsWith("/Session", StringComparison.OrdinalIgnoreCase))
                return Json("{\"SessionID\":\"test-session\"}");

            var query = request.RequestUri.Query;
            var owner = query.Contains("OwnerIDs", StringComparison.OrdinalIgnoreCase);
            var isNew = query.Contains("StateType=New", StringComparison.OrdinalIgnoreCase);
            return Json(owner
                ? isNew ? "{\"TicketID\":[3]}" : "{\"TicketID\":[1,2]}"
                : isNew ? "{\"TicketID\":[4]}" : "{\"TicketID\":[2]}" );
        }

        private static Task<HttpResponseMessage> Json(string json)
            => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(json, Encoding.UTF8, "application/json")
            });
    }
}
