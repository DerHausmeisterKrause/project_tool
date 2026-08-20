namespace TaskTool.Models;

public class WorkDayRecord
{
    public string Day { get; set; } = string.Empty;
    public DateTime? ComeLocal { get; set; }
    public DateTime? GoLocal { get; set; }
    public string DayType { get; set; } = "Normal"; // Normal, AM, UL
    public bool IsBr { get; set; }
    public bool IsHo { get; set; }
    public string HomeOfficeOutlookEntryId { get; set; } = string.Empty;
}

public sealed record SyncedCalendarMarkers(string Day, string OutlookDayType, bool OutlookIsHo);

public sealed record EffectiveDayMarkers(string DayType, bool IsHo);
