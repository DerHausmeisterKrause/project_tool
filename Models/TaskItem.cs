namespace TaskTool.Models;

public class TaskItem
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string Title { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string TicketUrl { get; set; } = string.Empty;
    public DateTime? StartLocal { get; set; }
    public DateTime? EndLocal { get; set; }
    public TaskStatus Status { get; set; } = TaskStatus.Planned;
    public int? Priority { get; set; }
    public string Tags { get; set; } = string.Empty;
    public string OutlookEntryId { get; set; } = string.Empty;
    public int TicketMinutesBooked { get; set; }
    public DateTime CreatedUtc { get; set; } = DateTime.UtcNow;
    public DateTime UpdatedUtc { get; set; } = DateTime.UtcNow;
}
