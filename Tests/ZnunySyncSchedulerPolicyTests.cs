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
}
