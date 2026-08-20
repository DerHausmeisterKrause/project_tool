namespace TaskTool.Models;

public sealed record TicketReplyTemplateContext(
    string Customer,
    string CustomerFirstName,
    string CustomerLastName,
    string TicketNumber,
    string TicketTitle,
    DateTime LocalDate);

public sealed record TicketReplyTemplateResult(string Text, bool HasUnresolvedVariables);
