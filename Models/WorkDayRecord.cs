namespace TaskTool.Models;

public class WorkDayRecord
{
    public string Day { get; set; } = string.Empty;
    public DateTime? ComeLocal { get; set; }
    public DateTime? GoLocal { get; set; }
    public string DayType { get; set; } = "Normal"; // Normal, AM, UL
    public bool IsBr { get; set; }
    public bool IsHo { get; set; }
}
