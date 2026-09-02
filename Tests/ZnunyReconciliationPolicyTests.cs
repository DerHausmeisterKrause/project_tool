using TaskTool.Models;
using Xunit;

namespace TaskTool.Tests;

public sealed class ZnunyReconciliationPolicyTests
{
    [Fact]
    public void RemovalVerificationWithoutTaskDoesNotCreateTask()
        => Assert.False(ZnunyReconciliationPolicy.ShouldCreateTask(isCurrentlyAssigned: false));

    [Fact]
    public void NewAssignmentWithoutTaskUsesNormalCreatePath()
        => Assert.True(ZnunyReconciliationPolicy.ShouldCreateTask(isCurrentlyAssigned: true));
}
