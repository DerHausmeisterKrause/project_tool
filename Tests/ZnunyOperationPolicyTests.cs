using TaskTool.Models;
using Xunit;

namespace TaskTool.Tests;

public sealed class ZnunyOperationPolicyTests
{
    [Theory]
    [InlineData("TicketSearchOwner")]
    [InlineData("TicketSearchResponsible")]
    [InlineData("TicketSearchCandidateOwnerActive")]
    [InlineData("TicketSearchCandidateResponsibleActive")]
    [InlineData("TicketGetDetails")]
    [InlineData("CandidateTicketGetDetails")]
    [InlineData("DynamicFieldOptions")]
    public void SemanticReadsAllowOneSessionRecovery(string operation)
        => Assert.True(ZnunyOperationPolicy.IsIdempotentRead(operation));

    [Theory]
    [InlineData("TicketCreate")]
    [InlineData("TicketUpdateSelfAssignment")]
    [InlineData("TicketUpdateReply")]
    [InlineData("TicketUpdateTimeBooking")]
    [InlineData("TicketUpdateTimeBookingRetry")]
    public void WritesNeverAllowAutomaticSessionRetry(string operation)
        => Assert.False(ZnunyOperationPolicy.IsIdempotentRead(operation));

    [Theory]
    [InlineData(1, false)]
    [InlineData(2, true)]
    [InlineData(3, true)]
    [InlineData(4, true)]
    public void SessionRecoveryRequiresSessionCreateAndRetryCapacity(int remaining, bool expected)
        => Assert.Equal(expected, ZnunyOperationPolicy.CanRecoverSession(remaining));

    [Theory]
    [InlineData(1, 0)]
    [InlineData(2, 0)]
    [InlineData(3, 1)]
    public void CandidateCapacityReservesSessionCreateAndSingleRetry(int remaining, int expected)
        => Assert.Equal(expected, ZnunyOperationPolicy.CandidateEvaluationsThatFit(remaining));

    [Fact]
    public void CandidateBatchWithOneRecoveryNearLimitStaysWithinBudget()
    {
        var used = ZnunySyncPolicy.MaximumAutomaticRequestsPerSync - 5;
        var candidates = ZnunyOperationPolicy.CandidateEvaluationsThatFit(
            ZnunySyncPolicy.MaximumAutomaticRequestsPerSync - used);

        Assert.Equal(ZnunySyncPolicy.MaximumAutomaticRequestsPerSync, used + candidates + 2);
    }
}
