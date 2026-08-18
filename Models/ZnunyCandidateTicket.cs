namespace TaskTool.Models;

public sealed class ZnunyCandidateTicket
{
    public string TicketId { get; init; } = string.Empty;
    public string TicketNumber { get; init; } = string.Empty;
    public string Title { get; init; } = string.Empty;
    public string DescriptionPreview { get; init; } = string.Empty;
    public string Owner { get; init; } = string.Empty;
    public string Responsible { get; init; } = string.Empty;
    public string State { get; init; } = string.Empty;
    public string WebUrl { get; init; } = string.Empty;
    public string MatchedKeyword { get; init; } = string.Empty;
}
