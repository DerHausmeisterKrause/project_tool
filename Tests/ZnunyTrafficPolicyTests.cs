using TaskTool.Models;
using Xunit;

namespace TaskTool.Tests;

public sealed class ZnunyTrafficPolicyTests
{
    [Theory]
    [InlineData(1, 5)]
    [InlineData(5, 5)]
    [InlineData(15, 15)]
    [InlineData(0, 15)]
    public void AutomaticSyncIntervalNeverFallsBelowFiveMinutes(int configured, int expected)
        => Assert.Equal(expected, ZnunyTrafficPolicy.NormalizeSyncIntervalMinutes(configured));

    [Fact]
    public void SyncSelectionContainsCurrentAndOnlyImmediatelyMissingPreviousTickets()
    {
        var current = Enumerable.Range(1, 20).Select(id => id.ToString()).ToList();
        var previous = Enumerable.Range(11, 20).Select(id => id.ToString()).ToList();
        var historical = Enumerable.Range(1000, 500).Select(id => id.ToString()).ToHashSet();

        var selected = ZnunyTrafficPolicy.SelectAssignedSyncTicketIds(current, previous);

        Assert.Equal(30, selected.Count);
        Assert.Empty(selected.Where(historical.Contains));
    }

    [Fact]
    public void ConservativeAutomaticBudgetsRemainBounded()
    {
        Assert.Equal(30, ZnunyTrafficPolicy.AutomaticRequestsPerMinute);
        Assert.Equal(25, ZnunyTrafficPolicy.CandidateTicketDetailsPerRefresh);
    }
}
