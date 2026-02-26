namespace TaskTool.Models;

public class AppSettings
{
    public bool OutlookSyncEnabled { get; set; } = true;
    public string OutlookCategoryName { get; set; } = "FocusBlock";
    public int ReminderLeadMinutes { get; set; } = 2;
    public string DateTimeFormat { get; set; } = "yyyy-MM-dd HH:mm";

    public int MondayTargetMinutes { get; set; } = 480;
    public int TuesdayTargetMinutes { get; set; } = 480;
    public int WednesdayTargetMinutes { get; set; } = 480;
    public int ThursdayTargetMinutes { get; set; } = 480;
    public int FridayTargetMinutes { get; set; } = 300;
    public int SaturdayTargetMinutes { get; set; } = 0;
    public int SundayTargetMinutes { get; set; } = 0;

    public int GetTargetMinutes(DayOfWeek day) => day switch
    {
        DayOfWeek.Monday => MondayTargetMinutes,
        DayOfWeek.Tuesday => TuesdayTargetMinutes,
        DayOfWeek.Wednesday => WednesdayTargetMinutes,
        DayOfWeek.Thursday => ThursdayTargetMinutes,
        DayOfWeek.Friday => FridayTargetMinutes,
        DayOfWeek.Saturday => SaturdayTargetMinutes,
        DayOfWeek.Sunday => SundayTargetMinutes,
        _ => 0
    };
}
