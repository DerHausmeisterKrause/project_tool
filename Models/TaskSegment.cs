namespace TaskTool.Models;

public class TaskSegment
{
    public long Id { get; set; }
    public Guid TaskId { get; set; }
    public DateTime StartLocal { get; set; }
    public DateTime EndLocal { get; set; }
    public int PlannedMinutes { get; set; }
    public string OutlookEntryId { get; set; } = string.Empty;
}
