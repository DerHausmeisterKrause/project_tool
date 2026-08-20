namespace TaskTool.Models;

public sealed record SuccessfulBookingStatistics(
    long TodaySeconds,
    IReadOnlyDictionary<DateTime, long> SecondsByMonth);

public sealed record MonthlyWorkDayStats(
    DateTime Month,
    int HomeOfficeDays,
    int VacationDays,
    int AmDays);

public sealed class MonthlyReportItem
{
    public DateTime Month { get; init; }
    public string MonthText { get; init; } = string.Empty;
    public long TotalSeconds { get; init; }
    public long TotalMinutes { get; init; }
    public string MinutesText { get; init; } = string.Empty;
    public string HoursText { get; init; } = string.Empty;
    public int HomeOfficeDays { get; init; }
    public int VacationDays { get; init; }
    public int AmDays { get; init; }
    public string HomeOfficeText { get; init; } = string.Empty;
    public string VacationText { get; init; } = string.Empty;
    public string AmText { get; init; } = string.Empty;
    public bool IsCurrentMonth { get; init; }
}
