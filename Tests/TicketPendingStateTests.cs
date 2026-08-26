using TaskTool.Models;
using Xunit;

namespace TaskTool.Tests;

public sealed class TicketPendingStateTests
{
    private static readonly DateTime Now = new(2026, 8, 26, 12, 30, 0, DateTimeKind.Utc);

    [Theory]
    [InlineData("pending reminder")]
    [InlineData("PENDING AUTO")]
    public void FutureSupportedPendingStateIsActive(string stateType)
    {
        var task = Ticket(stateType, Now.AddMinutes(5));
        Assert.True(TicketPendingState.IsActive(task, Now));
        Assert.False(TicketPendingState.IsWakeCandidate(task, Now));
    }

    [Fact]
    public void MissingTimeOrNonPendingStateNeverWakes()
    {
        Assert.False(TicketPendingState.IsWakeCandidate(Ticket("pending reminder", null), Now));
        Assert.False(TicketPendingState.IsWakeCandidate(Ticket("open", Now.AddMinutes(-5)), Now));
    }

    [Fact]
    public void ExactAndPastTimeAreWakeCandidatesUntilHandled()
    {
        Assert.True(TicketPendingState.IsWakeCandidate(Ticket("pending reminder", Now), Now));
        Assert.True(TicketPendingState.IsWakeCandidate(Ticket("pending reminder", Now.AddMinutes(-1)), Now));
        var handled = Ticket("pending reminder", Now);
        handled.PendingWakeHandledForUtc = Now;
        Assert.False(TicketPendingState.IsWakeCandidate(handled, Now));
    }

    [Fact]
    public void SameWakeIgnoresTicksAndDateTimeKind()
    {
        var utc = Now.AddTicks(9876);
        var unspecified = DateTime.SpecifyKind(Now, DateTimeKind.Unspecified);
        Assert.True(TicketPendingState.IsSameWake(utc, unspecified));
        Assert.Equal(TicketPendingState.CreateWakeKey("123", utc), TicketPendingState.CreateWakeKey("123", unspecified));
    }

    [Fact]
    public void NotificationAndHandledStateAreIndependent()
    {
        var task = Ticket("pending reminder", Now);
        task.PendingWakeHandledForUtc = Now;
        Assert.True(TicketPendingState.WasHandledFor(task, Now));
        Assert.False(TicketPendingState.WasNotificationSentFor(task, Now));
    }

    [Fact]
    public void NewPendingPeriodIsNotCoveredByOldHandledPeriod()
    {
        var task = Ticket("pending reminder", Now.AddHours(1));
        task.PendingWakeHandledForUtc = Now;
        Assert.False(TicketPendingState.WasHandledFor(task, Now.AddHours(1)));
    }

    [Fact]
    public void RepeatedCandidateChecksStaySuppressedAfterHandled()
    {
        var task = Ticket("pending reminder", Now);
        task.PendingWakeHandledForUtc = Now;
        Assert.All(Enumerable.Range(0, 100), _ => Assert.False(TicketPendingState.IsWakeCandidate(task, Now.AddMinutes(1))));
    }

    private static TaskItem Ticket(string stateType, DateTime? until) => new()
    {
        Tags = "Znuny;ZnunyTicketID:123",
        TicketStateType = stateType,
        TicketPendingUntilUtc = until
    };
}
