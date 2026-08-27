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

    [Fact]
    public void TicketSearchLimitIsExplicitAndBounded()
    {
        var payload = new Dictionary<string, object?> { ["SessionID"] = "secret" };
        ZnunySyncPolicy.ApplyTicketSearchLimit(payload);
        Assert.Equal(100, payload["Limit"]);
    }

    [Fact]
    public void MetadataTicketGetDoesNotRequestArticles()
    {
        var options = ZnunySyncPolicy.TicketGetOptions(allArticles: false, dynamicFields: false);
        Assert.Equal("0", options["AllArticles"]);
        Assert.Equal("0", options["DynamicFields"]);
        Assert.DoesNotContain("ArticleLimit", options.Keys);
    }

    [Fact]
    public void EveryArticleTicketGetIsLimitedToLatestTwenty()
    {
        var options = ZnunySyncPolicy.TicketGetOptions(allArticles: true, dynamicFields: true);
        Assert.Equal("1", options["AllArticles"]);
        Assert.Equal("20", options["ArticleLimit"]);
        Assert.Equal("DESC", options["ArticleOrder"]);
    }
}
