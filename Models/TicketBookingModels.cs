namespace TaskTool.Models;

public sealed record TicketFieldOption(string Key, string DisplayText)
{
    public override string ToString() => DisplayText;
}

public sealed record TicketBookingContext(
    string TicketId,
    string TicketNumber,
    string CostCenterValue,
    string OrderValue,
    IReadOnlyList<TicketFieldOption> CostCenterOptions,
    IReadOnlyList<TicketFieldOption> OrderOptions,
    string Information);

public sealed record TicketBookingResult(bool Success, bool PendingReconciliation, string Message);
