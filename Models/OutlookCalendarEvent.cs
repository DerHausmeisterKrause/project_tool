namespace TaskTool.Models;

public class OutlookCalendarEvent
{
    public string Id { get; set; } = string.Empty;
    public string EntryId { get; set; } = string.Empty;
    public string ICalUId { get; set; } = string.Empty;
    public string CalendarName { get; set; } = string.Empty;
    public string BusyStatus { get; set; } = string.Empty;
    public string Sensitivity { get; set; } = string.Empty;
    public bool IsPrivate { get; set; }
    public bool IsRecurring { get; set; }
    public bool IsInstance { get; set; }
    public string Subject { get; set; } = string.Empty;
    public DateTime StartLocal { get; set; }
    public DateTime EndLocal { get; set; }
    public bool IsAllDay { get; set; }
    public string Location { get; set; } = string.Empty;
    public string Organizer { get; set; } = string.Empty;
    public string BodyPreview { get; set; } = string.Empty;
    public string OnlineMeetingJoinUrl { get; set; } = string.Empty;
    public string Categories { get; set; } = string.Empty;
    public string MeetingStatus { get; set; } = string.Empty;
    public string MessageClass { get; set; } = string.Empty;
    public bool IsCancelled { get; set; }

    public bool HasTeamsLink => !string.IsNullOrWhiteSpace(OnlineMeetingJoinUrl);
}
