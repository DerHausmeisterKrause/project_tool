namespace TaskTool.Models;

public class AppSettings
{
    public bool DynamicIslandEnabled { get; set; } = true;
    public string DynamicIslandDockPosition { get; set; } = "TopCenter";

    public bool OutlookSyncEnabled { get; set; } = true;
    public string OutlookCategoryName { get; set; } = "FocusBlock";

    public bool OutlookCalendarEnabled { get; set; } = false;
    public bool OutlookConflictWarningsEnabled { get; set; } = true;
    public bool OutlookTeamsButtonEnabled { get; set; } = true;
    public string OutlookCalendarSyncMode { get; set; } = "Manual";
    public int OutlookCalendarSyncIntervalMinutes { get; set; } = 5;
    public int OutlookCalendarRangePastDays { get; set; } = 0;
    public int OutlookCalendarRangeFutureDays { get; set; } = 14;
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
