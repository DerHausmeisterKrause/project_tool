namespace TaskTool.Models;

public sealed record AssignTicketResult(
    bool Success,
    string Message,
    bool ConfirmationUncertain = false);
