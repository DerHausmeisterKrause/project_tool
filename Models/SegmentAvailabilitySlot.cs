namespace TaskTool.Models;

public sealed class SegmentAvailabilitySlot
{
    public required DateTime Start { get; init; }
    public required DateTime End { get; init; }
    public bool IsBusy { get; init; }

    public string TimeText => $"{Start:HH:mm}–{End:HH:mm}";
}
