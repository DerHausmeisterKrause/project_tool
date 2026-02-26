namespace TaskTool.Models;

public class TimeLog
{
    public long Id { get; set; }
    public Guid TaskId { get; set; }
    public DateTime StartUtc { get; set; }
    public DateTime? EndUtc { get; set; }
    public string Note { get; set; } = string.Empty;
}
