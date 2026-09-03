using TaskTool.Models;

namespace TaskTool.ViewModels;

public sealed class TodayAgendaItem
{
    public DateTime Start { get; init; }
    public DateTime End { get; init; }
    public string Title { get; init; } = string.Empty;
    public string Location { get; init; } = string.Empty;
    public TaskItem? Task { get; init; }
    public TaskSegment? Segment { get; init; }
    public OutlookCalendarEvent? OutlookEvent { get; init; }
    public bool IsAllDay { get; init; }
    public bool IsTask => Task != null;
    public bool IsOutlook => OutlookEvent != null;
    public bool HasTeamsLink => OutlookEvent?.HasTeamsLink == true;
    public bool HasUnreadTicketArticles => Task?.HasUnreadTicketArticles == true;
    public Guid TaskId => Task?.Id ?? Guid.Empty;
    public string TypeLabel => IsTask ? "Task" : "Outlook";
    public string TimeLabel => IsAllDay ? "Ganztägig" : $"{Start:HH:mm} – {End:HH:mm}";
}
