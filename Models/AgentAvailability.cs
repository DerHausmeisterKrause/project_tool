namespace TaskTool.Models;

public enum AgentAvailabilityLevel
{
    Unknown,
    Green,
    Orange,
    Red
}

public sealed record AgentAvailabilitySnapshot(
    AgentAvailabilityLevel Level,
    DateTime CalculatedUtc);

internal sealed record OutlookFreeBusySnapshot(
    DateTime StartLocal,
    int IntervalMinutes,
    string Slots);
