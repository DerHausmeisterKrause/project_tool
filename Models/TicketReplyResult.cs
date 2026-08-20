namespace TaskTool.Models;

public sealed record TicketReplyResult(
    bool Success,
    string Message,
    string ArticleId = "",
    bool ConfirmationUncertain = false);
