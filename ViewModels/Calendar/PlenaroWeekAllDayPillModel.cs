namespace TaskTool.ViewModels;

public class PlenaroWeekAllDayPillModel
{
    public string Id { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public string Location { get; set; } = string.Empty;
    public string TeamsJoinUrl { get; set; } = string.Empty;
    public bool HasTeamsLink => !string.IsNullOrWhiteSpace(TeamsJoinUrl);
}
