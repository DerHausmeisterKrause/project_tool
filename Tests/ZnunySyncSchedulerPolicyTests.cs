using TaskTool.Models;
using Xunit;

namespace TaskTool.Tests;

public sealed class ZnunySyncSchedulerPolicyTests
{
    [Fact]
    public void SlotsDoNotDriftWithRunDuration()
    {
        var anchor = new DateTimeOffset(2026, 1, 1, 10, 0, 0, TimeSpan.Zero);
        Assert.Equal(anchor.AddMinutes(15), ZnunySyncSchedulerPolicy.NextSlot(anchor, anchor.AddMinutes(2), TimeSpan.FromMinutes(15)));
        Assert.Equal(anchor.AddMinutes(30), ZnunySyncSchedulerPolicy.NextSlot(anchor, anchor.AddMinutes(17), TimeSpan.FromMinutes(15)));
    }

    [Fact]
    public void StartupOffsetsAreStableBoundedAndUseDifferentSalts()
    {
        const string id = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee";
        var candidate = ZnunySyncSchedulerPolicy.StartupCandidateDelay(id);
        var full = ZnunySyncSchedulerPolicy.StartupFullDelay(id);
        Assert.Equal(candidate, ZnunySyncSchedulerPolicy.StartupCandidateDelay(id));
        Assert.InRange(candidate.TotalSeconds, 2, 12);
        Assert.InRange(full.TotalSeconds, 5, 30);
        Assert.NotEqual(candidate, full);
    }

    [Fact]
    public void CandidateBusyUsesShortRetryButFailuresUseRegularSlot()
    {
        const string id = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee";
        var anchor = new DateTimeOffset(2026, 1, 1, 10, 0, 7, TimeSpan.Zero);
        var now = anchor.AddMinutes(1);
        var interval = TimeSpan.FromMinutes(5);

        var busy = ZnunySyncSchedulerPolicy.DelayAfterRun(id, "candidate", ZnunyScheduledRunOutcome.Busy, anchor, now, interval);
        var apiFailure = ZnunySyncSchedulerPolicy.DelayAfterRun(id, "candidate", ZnunyScheduledRunOutcome.Failed, anchor, now, interval);
        var configFailure = ZnunySyncSchedulerPolicy.DelayAfterRun(id, "candidate", ZnunyScheduledRunOutcome.Failed, anchor, now, interval);

        Assert.InRange(busy.TotalSeconds, 10, 30);
        Assert.Equal(TimeSpan.FromMinutes(4), apiFailure);
        Assert.Equal(TimeSpan.FromMinutes(4), configFailure);
    }

    [Fact]
    public void StartupFullBecomesAnchorForDriftFreeRegularIntervals()
    {
        const string id = "11111111-2222-3333-4444-555555555555";
        var started = new DateTimeOffset(2026, 1, 1, 10, 0, 0, TimeSpan.Zero);
        var fullDue = started + ZnunySyncSchedulerPolicy.StartupFullDelay(id);

        Assert.InRange((fullDue - started).TotalSeconds, 5, 30);
        Assert.Equal(fullDue.AddMinutes(15), ZnunySyncSchedulerPolicy.NextSlot(fullDue, fullDue.AddSeconds(2), TimeSpan.FromMinutes(15)));
        Assert.Equal(fullDue.AddMinutes(30), ZnunySyncSchedulerPolicy.NextSlot(fullDue, fullDue.AddMinutes(17), TimeSpan.FromMinutes(15)));
    }
}
