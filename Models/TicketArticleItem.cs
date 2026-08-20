namespace TaskTool.Models;

public sealed class TicketArticleItem
{
    public string ArticleId { get; init; } = string.Empty;
    public DateTime? CreatedLocal { get; init; }
    public string Subject { get; init; } = string.Empty;
    public string Body { get; init; } = string.Empty;
    public string SenderType { get; init; } = string.Empty;
    public string CommunicationChannel { get; init; } = string.Empty;
    public string From { get; init; } = string.Empty;
    public string To { get; init; } = string.Empty;
    public string ReplyTo { get; init; } = string.Empty;
    public string MessageId { get; init; } = string.Empty;
    public bool IsVisibleForCustomer { get; init; }
    public string TypeText { get; init; } = string.Empty;
    public string DisplayText { get; init; } = string.Empty;
    public string CreatedDisplay => CreatedLocal?.ToString("dd.MM.yyyy HH:mm") ?? "Zeit unbekannt";
    public string FromDisplay => string.IsNullOrWhiteSpace(From) ? "Unbekannt" : From;
}
