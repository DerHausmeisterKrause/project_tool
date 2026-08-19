using System.Collections.ObjectModel;
using System.Globalization;
using TaskTool.Infrastructure;
using TaskTool.Models;
using TaskTool.Services;

namespace TaskTool.ViewModels;

public class ReportsViewModel : ObservableObject
{
    private static readonly CultureInfo ReportCulture = CultureInfo.GetCultureInfo("de-DE");
    private readonly TaskService _tasks;
    private readonly WorkDayService _workDays;
    private readonly SettingsService _settings;
    private readonly GermanTimeService _germanTime;
    private readonly LoggerService _logger;

    public string Title => "Reports";

    private long _bookedSecondsToday;
    public long BookedSecondsToday { get => _bookedSecondsToday; private set => Set(ref _bookedSecondsToday, value); }

    private string _bookedMinutesTodayText = "0 Min.";
    public string BookedMinutesTodayText { get => _bookedMinutesTodayText; private set => Set(ref _bookedMinutesTodayText, value); }

    private string _bookedHoursTodayText = "0 Std. 00 Min.";
    public string BookedHoursTodayText { get => _bookedHoursTodayText; private set => Set(ref _bookedHoursTodayText, value); }

    public ObservableCollection<MonthlyReportItem> MonthlyReports { get; } = new();
    public RelayCommand RefreshCommand { get; }

    public ReportsViewModel(
        TaskService tasks,
        WorkDayService workDays,
        SettingsService settings,
        GermanTimeService germanTime,
        LoggerService logger)
    {
        _tasks = tasks;
        _workDays = workDays;
        _settings = settings;
        _germanTime = germanTime;
        _logger = logger;
        RefreshCommand = new RelayCommand(Refresh);
        Refresh();
    }

    public void Refresh()
    {
        try
        {
            var timeZone = _germanTime.ResolveTimeZone(_settings.Current.CalendarTimeZoneId);
            var localToday = _germanTime.GetLocalNow(_settings.Current.CalendarTimeZoneId).Date;
            var currentMonth = new DateTime(localToday.Year, localToday.Month, 1);
            var bookingStats = _tasks.GetSuccessfulBookingStatistics(localToday, timeZone);
            var markerStats = _workDays.GetMonthlyMarkerStatistics()
                .ToDictionary(item => item.Month);

            BookedSecondsToday = bookingStats.TodaySeconds;
            var todayMinutes = WholeMinutes(BookedSecondsToday);
            BookedMinutesTodayText = FormatMinutes(todayMinutes);
            BookedHoursTodayText = FormatHours(todayMinutes);

            var months = bookingStats.SecondsByMonth.Keys
                .Concat(markerStats.Keys)
                .Append(currentMonth)
                .Where(month => month <= currentMonth)
                .Distinct()
                .OrderByDescending(month => month)
                .ToList();

            var reports = months.Select(month =>
            {
                var seconds = bookingStats.SecondsByMonth.GetValueOrDefault(month);
                var minutes = WholeMinutes(seconds);
                markerStats.TryGetValue(month, out var markers);
                var homeOfficeDays = markers?.HomeOfficeDays ?? 0;
                var vacationDays = markers?.VacationDays ?? 0;
                var amDays = markers?.AmDays ?? 0;
                return new MonthlyReportItem
                {
                    Month = month,
                    MonthText = month.ToString("MMMM yyyy", ReportCulture),
                    TotalSeconds = seconds,
                    TotalMinutes = minutes,
                    MinutesText = FormatMinutes(minutes),
                    HoursText = FormatHours(minutes),
                    HomeOfficeDays = homeOfficeDays,
                    VacationDays = vacationDays,
                    AmDays = amDays,
                    HomeOfficeText = FormatDays(homeOfficeDays),
                    VacationText = FormatDays(vacationDays),
                    AmText = FormatDays(amDays),
                    IsCurrentMonth = month == currentMonth
                };
            }).ToList();

            MonthlyReports.Clear();
            foreach (var report in reports)
                MonthlyReports.Add(report);

            var current = reports.First(item => item.IsCurrentMonth);
            _logger.Info($"[Reports] todaySeconds={BookedSecondsToday} months={reports.Count} currentMonthSeconds={current.TotalSeconds} currentMonthHo={current.HomeOfficeDays} currentMonthUl={current.VacationDays} currentMonthAm={current.AmDays}");
        }
        catch (Exception ex)
        {
            _logger.Error($"[Reports] refreshFailed=true message='{ex.Message}'");
            BookedSecondsToday = 0;
            BookedMinutesTodayText = FormatMinutes(0);
            BookedHoursTodayText = FormatHours(0);
            MonthlyReports.Clear();
        }
    }

    private static long WholeMinutes(long seconds) => Math.Max(0, seconds) / 60;
    private static string FormatMinutes(long minutes) => $"{minutes.ToString("N0", ReportCulture)} Min.";
    private static string FormatHours(long minutes) => $"{minutes / 60} Std. {minutes % 60:00} Min.";
    private static string FormatDays(int days) => $"{days} {(days == 1 ? "Tag" : "Tage")}";

    public override string ToString() => Title;
}
