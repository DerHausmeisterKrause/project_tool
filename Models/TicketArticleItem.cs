using System.Text.Json.Serialization;
using TaskTool.Infrastructure;

namespace TaskTool.Models;

public sealed class TicketArticleItem : ObservableObject
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
    [JsonIgnore]
    private bool _isUnread;
    public bool IsUnread
    {
        get => _isUnread;
        set { if (Set(ref _isUnread, value)) Raise(nameof(DropdownDisplayText)); }
    }
    [JsonIgnore]
    public string DropdownDisplayText => IsUnread ? $"★ {DisplayText}" : DisplayText;
    public string CreatedDisplay => CreatedLocal?.ToString("dd.MM.yyyy HH:mm") ?? "Zeit unbekannt";
    public string FromDisplay => string.IsNullOrWhiteSpace(From) ? "Unbekannt" : From;
}
