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
        ZnunySyncPolicy.ApplyTicketSearchLimit(payload, 100);
        Assert.Equal(100, payload["Limit"]);
    }

    [Fact]
    public void MetadataTicketGetDoesNotRequestArticles()
    {
        var options = ZnunySyncPolicy.TicketGetOptions(allArticles: false, dynamicFields: false, configuredArticleLimit: 17);
        Assert.Equal("0", options["AllArticles"]);
        Assert.Equal("0", options["DynamicFields"]);
        Assert.DoesNotContain("ArticleLimit", options.Keys);
    }

    [Fact]
    public void EveryArticleTicketGetIsLimitedToLatestTwenty()
    {
        var options = ZnunySyncPolicy.TicketGetOptions(allArticles: true, dynamicFields: true, configuredArticleLimit: 20);
        Assert.Equal("1", options["AllArticles"]);
        Assert.Equal("20", options["ArticleLimit"]);
        Assert.Equal("DESC", options["ArticleOrder"]);
    }

    [Theory]
    [InlineData(0, 10)]
    [InlineData(10, 10)]
    [InlineData(100, 100)]
    [InlineData(500, 500)]
    [InlineData(501, 500)]
    [InlineData(10000, 500)]
    public void SearchLimitIsNormalized(int configured, int expected)
        => Assert.Equal(expected, ZnunySyncPolicy.NormalizeSearchLimit(configured));

    [Theory]
    [InlineData(0, 1)]
    [InlineData(1, 1)]
    [InlineData(20, 20)]
    [InlineData(100, 100)]
    [InlineData(101, 100)]
    [InlineData(999, 100)]
    public void ArticleLimitIsNormalized(int configured, int expected)
        => Assert.Equal(expected, ZnunySyncPolicy.NormalizeArticleLimit(configured));

    [Fact]
    public void ConfiguredLimitsAreAppliedToRequests()
    {
        var payload = new Dictionary<string, object?>();
        ZnunySyncPolicy.ApplyTicketSearchLimit(payload, 73);
        var options = ZnunySyncPolicy.TicketGetOptions(true, false, 17);

        Assert.Equal(73, payload["Limit"]);
        Assert.Equal("1", options["AllArticles"]);
        Assert.Equal("DESC", options["ArticleOrder"]);
        Assert.Equal("17", options["ArticleLimit"]);
    }

    [Theory]
    [InlineData(0, 5)]
    [InlineData(1, 3)]
    [InlineData(2, 3)]
    [InlineData(3, 3)]
    [InlineData(5, 5)]
    [InlineData(60, 60)]
    [InlineData(61, 60)]
    [InlineData(500, 60)]
    public void CandidateIntervalIsNormalized(int configured, int expected)
        => Assert.Equal(expected, ZnunySyncPolicy.NormalizeCandidateIntervalMinutes(configured));

    [Fact]
    public void CandidateTimerIsAnAutomaticRequestReason()
        => Assert.True(ZnunyRequestReason.CandidateTimerSync.IsAutomatic());

    [Theory]
    [InlineData("Owner", "OwnerIDs")]
    [InlineData("Responsible", "ResponsibleIDs")]
    public void AssignedRoleIdsAreArraysAndSearchIsExplicitlySorted(string role, string key)
    {
        var payload = new Dictionary<string, object?>();
        ZnunySyncPolicy.ApplyTicketRoleCriteria(payload, role, 123, onlyOpen: true);
        ZnunySyncPolicy.ApplyTicketSearchLimit(payload, 500);

        Assert.Equal(new[] { 123 }, Assert.IsType<int[]>(payload[key]));
        Assert.Equal(new[] { "new", "open" }, Assert.IsType<string[]>(payload["StateType"]));
        Assert.Equal(500, payload["Limit"]);
        Assert.Equal("Changed", payload["SortBy"]);
        Assert.Equal("Down", payload["OrderBy"]);
    }

    [Fact]
    public void BusySyncIsNeverSuccessful()
    {
        var result = ZnunySyncResult.BusyResult();
        Assert.False(result.Started);
        Assert.False(result.Success);
        Assert.True(result.Busy);
    }

    [Fact]
    public void GenuineZeroResultCanBeSuccessful()
    {
        var result = new ZnunySyncResult(true, true, false, 0, 0, 0, 0, 0, 0, false, string.Empty);
        Assert.True(result.Started);
        Assert.True(result.Success);
        Assert.False(result.Busy);
    }

    [Fact]
    public void FailedSyncIsNeverSuccessful()
    {
        var result = ZnunySyncResult.Failed("HTTP 500");
        Assert.True(result.Started);
        Assert.False(result.Success);
        Assert.Equal("HTTP 500", result.ErrorMessage);
    }

    [Fact]
    public void SuccessfulResultKeepsUnchangedCountSeparate()
    {
        var result = new ZnunySyncResult(true, true, false, 87, 84, 1, 2, 81, 0, false, string.Empty);
        Assert.Equal(81, result.Unchanged);
        Assert.Equal(84, result.Created + result.Updated + result.Unchanged + result.Skipped);
    }
}
