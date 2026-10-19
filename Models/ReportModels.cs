namespace TaskTool.Models;

public sealed record SuccessfulBookingStatistics(
    long TodaySeconds,
    IReadOnlyDictionary<DateTime, long> SecondsByMonth);

public sealed record MonthlyWorkDayStats(
    DateTime Month,
    int HomeOfficeDays,
    int VacationDays,
    int AmDays);

public sealed class MonthlyReportTimeAdjustment
{
    public Guid Id { get; init; } = Guid.NewGuid();
    public DateTime Month { get; init; }
    public long Seconds { get; init; }
    public string Note { get; init; } = string.Empty;
    public DateTime CreatedUtc { get; init; } = DateTime.UtcNow;
    public string DurationText => $"{Seconds / 3600:00}:{Seconds % 3600 / 60:00} Std.";
    public string SummaryText => string.IsNullOrWhiteSpace(Note) ? DurationText : $"{DurationText} · {Note}";
    public string CreatedText => CreatedUtc.ToLocalTime().ToString("dd.MM.yyyy");
}

public sealed class MonthlyReportItem
{
    public DateTime Month { get; init; }
    public string MonthText { get; init; } = string.Empty;
    public long TotalSeconds { get; init; }
    public long BookedSeconds { get; init; }
    public long ManualAdjustmentSeconds { get; init; }
    public long TotalMinutes { get; init; }
    public string MinutesText { get; init; } = string.Empty;
    public string HoursText { get; init; } = string.Empty;
    public string ManualAdjustmentText { get; init; } = string.Empty;
    public bool HasManualAdjustments => ManualAdjustmentSeconds > 0;
    public int HomeOfficeDays { get; init; }
    public int VacationDays { get; init; }
    public int AmDays { get; init; }
    public string HomeOfficeText { get; init; } = string.Empty;
    public string VacationText { get; init; } = string.Empty;
    public string AmText { get; init; } = string.Empty;
    public bool IsCurrentMonth { get; init; }
}
