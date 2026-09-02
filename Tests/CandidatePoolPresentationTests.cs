using TaskTool.Models;
using Xunit;

namespace TaskTool.Tests;

public sealed class CandidatePoolPresentationTests
{
    [Fact]
    public void ShowAllPoolPreferenceDefaultsOff()
        => Assert.False(new AppSettings().ShowAllCandidatePoolTickets);

    [Fact]
    public void MatchedAndUnmatchedTicketsRetainDistinctPresentationState()
    {
        var matched = new ZnunyCandidateTicket { TicketId = "1", MatchedKeyword = "server" };
        var unmatched = new ZnunyCandidateTicket { TicketId = "2", MatchedKeyword = "" };
        Assert.True(matched.HasMatchedKeyword);
        Assert.False(unmatched.HasMatchedKeyword);
    }

    [Fact]
    public void LocalToggleSelectsFilteredOrWholePoolWithoutAnyRefreshDependency()
    {
        var matched = new ZnunyCandidateTicket { TicketId = "1", MatchedKeyword = "server" };
        var unmatched = new ZnunyCandidateTicket { TicketId = "2" };
        IReadOnlyList<ZnunyCandidateTicket> filtered = [matched];
        IReadOnlyList<ZnunyCandidateTicket> pool = [matched, unmatched];

        Assert.Equal([matched], CandidateTicketViewPolicy.Select(filtered, pool, showAll: false));
        Assert.Equal([matched, unmatched], CandidateTicketViewPolicy.Select(filtered, pool, showAll: true));
    }
}
