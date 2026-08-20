namespace TaskTool.Models;

public sealed record CreateTicketResult(
    bool Success,
    string Message,
    string TicketId = "",
    string TicketNumber = "",
    string TicketUrl = "",
    bool ConfirmationUncertain = false);
