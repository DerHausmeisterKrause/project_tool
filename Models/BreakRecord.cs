namespace TaskTool.Models;

public class BreakRecord
{
    public long Id { get; set; }
    public string Day { get; set; } = string.Empty;
    public DateTime StartLocal { get; set; }
    public DateTime? EndLocal { get; set; }
    public string Note { get; set; } = string.Empty;
}
