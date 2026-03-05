namespace TaskTool.Models;

public class OutlookCalendarEvent
{
    public string Id { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public DateTime StartLocal { get; set; }
    public DateTime EndLocal { get; set; }
    public bool IsAllDay { get; set; }
    public string Location { get; set; } = string.Empty;
    public string Organizer { get; set; } = string.Empty;
    public string BodyPreview { get; set; } = string.Empty;
    public string OnlineMeetingJoinUrl { get; set; } = string.Empty;
    public string Categories { get; set; } = string.Empty;

    public bool HasTeamsLink => !string.IsNullOrWhiteSpace(OnlineMeetingJoinUrl);
}
