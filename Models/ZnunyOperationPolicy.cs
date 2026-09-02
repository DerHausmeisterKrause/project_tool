namespace TaskTool.Models;

public static class ZnunyOperationPolicy
{
    private static readonly HashSet<string> ExactReadOperations = new(StringComparer.Ordinal)
    {
        "TicketSearchOwner", "TicketSearchResponsible",
        "TicketSearchCandidateOwnerActive", "TicketSearchCandidateResponsibleActive",
        "DynamicFieldOptions", "SessionGetDiagnostic"
    };

    public static bool IsIdempotentRead(string operation)
        => ExactReadOperations.Contains(operation)
           || operation.StartsWith("TicketGet", StringComparison.Ordinal)
           || operation.StartsWith("CandidateTicketGet", StringComparison.Ordinal);

}
