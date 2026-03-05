namespace TaskTool.ViewModels;

public class PlenaroWeekOutlookEventBlock
{
    public string Id { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public string TimeLabel { get; set; } = string.Empty;
    public string Location { get; set; } = string.Empty;
    public string TeamsJoinUrl { get; set; } = string.Empty;
    public DateTime StartLocal { get; set; }
    public DateTime EndLocal { get; set; }
    public double DisplayTop { get; set; }
    public double DisplayHeight { get; set; }
    public double DisplayLeft { get; set; }
    public double DisplayWidth { get; set; }
    public int OverlapColumn { get; set; }
    public int OverlapColumnCount { get; set; }
    public bool IsCompact { get; set; }
    public bool ShowLocation { get; set; }
    public bool ShowActions { get; set; }
    public bool HasTeamsLink => !string.IsNullOrWhiteSpace(TeamsJoinUrl);
    public string TooltipText { get; set; } = string.Empty;
}
