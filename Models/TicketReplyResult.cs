namespace TaskTool.Models;

public enum TicketReplyMode
{
    Normal,
    CloseTicket
}

public sealed record TicketReplyResult(
    bool Success,
    string Message,
    string ArticleId = "",
    bool ConfirmationUncertain = false,
    bool TicketClosed = false);
