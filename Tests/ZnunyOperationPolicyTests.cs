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

}
