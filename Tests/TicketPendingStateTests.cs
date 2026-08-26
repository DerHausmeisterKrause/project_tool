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

    [Theory]
    [InlineData(7200, 14, 0)]
    [InlineData(300, 12, 5)]
    [InlineData(-30, 11, 59)]
    public void UntilTimeIsSignedDurationNotUnixTimestamp(long seconds, int expectedHour, int expectedMinute)
    {
        var responseReceivedUtc = new DateTime(2026, 8, 26, 12, 0, 0, DateTimeKind.Utc);
        var resolved = TicketPendingState.ResolveRelativePendingUtc(responseReceivedUtc, seconds);
        Assert.True(resolved.HasValue);
        var value = resolved.GetValueOrDefault();
        Assert.Equal(expectedHour, value.Hour);
        Assert.Equal(expectedMinute, value.Minute);
        Assert.Equal(2026, value.Year);
    }

    [Fact]
    public void Legacy1970PendingValueNeverWakes()
    {
        Assert.False(TicketPendingState.IsWakeCandidate(
            Ticket("pending reminder", new DateTime(1970, 1, 1, 2, 0, 0, DateTimeKind.Utc)), Now));
    }

    [Fact]
    public void StaleUnassignedZnunyTaskRemainsVisibleUntilInitialSyncCompletes()
    {
        var task = Ticket("open", null);
        task.IsZnunyAssigned = false;

        Assert.True(task.IsOperationallyVisibleWithRemoteState(canTrustRemoteTicketState: false));
        Assert.False(task.IsOperationallyVisibleWithRemoteState(canTrustRemoteTicketState: true));
    }

    [Fact]
    public void LocalTaskVisibilityNeverDependsOnZnunySync()
    {
        var task = new TaskItem { IsZnunyAssigned = false };

        Assert.True(task.IsOperationallyVisibleWithRemoteState(canTrustRemoteTicketState: false));
        Assert.True(task.IsOperationallyVisibleWithRemoteState(canTrustRemoteTicketState: true));
    }

    private static TaskItem Ticket(string stateType, DateTime? until) => new()
    {
        Tags = "Znuny;ZnunyTicketID:123",
        TicketStateType = stateType,
        TicketPendingUntilUtc = until
    };
}
