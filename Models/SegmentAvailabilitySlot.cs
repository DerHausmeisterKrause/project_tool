namespace TaskTool.Models;

public sealed class SegmentAvailabilitySlot
{
    public required DateTime Start { get; init; }
    public required DateTime End { get; init; }
    public bool IsBusy { get; init; }
    public bool IsUnknown { get; init; }
    public bool IsAvailable => !IsUnknown && !IsBusy;

    public string TimeText => $"{Start:HH:mm}–{End:HH:mm}";
    public string Tooltip => IsUnknown ? $"{TimeText}: Kalenderdaten nicht verfügbar"
        : IsBusy ? $"{TimeText}: durch Outlook belegt" : $"{TimeText}: frei";

    public bool Overlaps(DateTime appointmentStart, DateTime appointmentEnd)
        => appointmentStart < End && appointmentEnd > Start;
}
