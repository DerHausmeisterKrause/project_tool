using TaskTool.Models;
using Xunit;

namespace TaskTool.Tests;

public sealed class WikiSearchPolicyTests
{
    [Fact]
    public void ManualRefreshAllowsLocallyStoredUnassignedZnunyTicket()
    {
        var task = ZnunyTask(assigned: false);

        Assert.True(WikiSearchPolicy.CanSearch(task, force: true));
    }

    [Fact]
    public void AutomaticSearchRemainsRestrictedToAssignedTickets()
    {
        var task = ZnunyTask(assigned: false);

        Assert.False(WikiSearchPolicy.CanSearch(task, force: false));
    }

    [Fact]
    public void NonZnunyTaskCannotStartManualWikiSearch()
        => Assert.False(WikiSearchPolicy.CanSearch(new TaskItem(), force: true));

    [Fact]
    public void TaskValuesAreUsedWhenLocalDetailCacheIsEmpty()
    {
        var task = ZnunyTask(assigned: false);
        task.Title = "Lokaler Titel";
        task.Description = "Lokale Beschreibung";

        Assert.Equal("Lokaler Titel", WikiSearchPolicy.ResolveTitle(task, string.Empty));
        Assert.Equal("Lokale Beschreibung", WikiSearchPolicy.ResolveMessage(task, null));
    }

    [Fact]
    public void LocalDetailCacheValuesTakePrecedence()
    {
        var task = ZnunyTask(assigned: true);

        Assert.Equal("Cache-Titel", WikiSearchPolicy.ResolveTitle(task, "Cache-Titel"));
        Assert.Equal("Cache-Nachricht", WikiSearchPolicy.ResolveMessage(task, "Cache-Nachricht"));
    }

    private static TaskItem ZnunyTask(bool assigned) => new()
    {
        Tags = "ZnunyTicketID:123",
        IsZnunyAssigned = assigned
    };
}
