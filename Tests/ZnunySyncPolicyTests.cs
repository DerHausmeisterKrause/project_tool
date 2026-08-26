using TaskTool.Models;
using Xunit;

namespace TaskTool.Tests;

public sealed class ZnunySyncPolicyTests
{
    [Theory]
    [InlineData(1, 5)]
    [InlineData(5, 5)]
    [InlineData(15, 15)]
    [InlineData(0, 15)]
    public void IntervalIsNeverBelowFiveMinutes(int configured, int expected)
        => Assert.Equal(expected, ZnunySyncPolicy.NormalizeIntervalMinutes(configured));

    [Fact]
    public void HistoricalLocalTicketsAreNotPartOfRemoteSelection()
    {
        var current = Enumerable.Range(1, 15).Select(value => value.ToString()).ToList();
        var previous = current.Concat(new[] { "16" }).ToList();
        var selected = ZnunySyncPolicy.SelectTicketIds(current, previous);

        Assert.Equal(16, selected.Count);
        Assert.DoesNotContain("500", selected);
    }

    [Fact]
    public void AutomaticRequestBudgetIsHardBounded()
        => Assert.Equal(60, ZnunySyncPolicy.MaximumAutomaticRequestsPerSync);
}
