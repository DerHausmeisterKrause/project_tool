using System.Collections.ObjectModel;
using TaskTool.Infrastructure;
using TaskTool.Services;

namespace TaskTool.ViewModels;

public class ReportsViewModel : ObservableObject
{
    private readonly TaskService _tasks;
    private readonly WorkDayService _workDays;
    private readonly SettingsService _settings;

    public string Title => "Reports";

    private string _monthSummary = "0h 0m";
    public string MonthSummary { get => _monthSummary; set => Set(ref _monthSummary, value); }

    public string NetMonthSummary { get => _netMonthSummary; set => Set(ref _netMonthSummary, value); }
    private string _netMonthSummary = "0h 0m";
    public string SollMonthSummary { get => _sollMonthSummary; set => Set(ref _sollMonthSummary, value); }
    private string _sollMonthSummary = "0h 0m";
    public string OvertimeMonthSummary { get => _overtimeMonthSummary; set => Set(ref _overtimeMonthSummary, value); }
    private string _overtimeMonthSummary = "0h 0m";

    public int TotalMinutes { get => _totalMinutes; set => Set(ref _totalMinutes, value); }
    private int _totalMinutes;

    public ObservableCollection<string> Breakdown { get; } = new();
    public ObservableCollection<ReportTaskItem> TopTasks { get; } = new();
    public RelayCommand RefreshCommand { get; }

    public ReportsViewModel(TaskService tasks, WorkDayService workDays, SettingsService settings)
    {
        _tasks = tasks;
        _workDays = workDays;
        _settings = settings;
        RefreshCommand = new RelayCommand(Refresh);
        Refresh();
    }

    private void Refresh()
    {
        try
        {
            var month = DateTime.Today;
            TotalMinutes = _tasks.GetMonthTicketMinutes(month);
            MonthSummary = Fmt(TotalMinutes);

            var first = new DateTime(month.Year, month.Month, 1);
            var last = first.AddMonths(1).AddDays(-1);
            var days = _workDays.GetWorkDaysInRange(first, last).ToDictionary(d => d.Day, d => d);
            var totalNet = 0;
            var totalSoll = 0;

            for (var day = first; day <= last; day = day.AddDays(1))
            {
                var key = day.ToString("yyyy-MM-dd");
                var wd = days.ContainsKey(key) ? days[key] : new Models.WorkDayRecord { Day = key, DayType = "Normal" };
                var target = (wd.DayType == "UL" || wd.DayType == "AM") ? 0 : _settings.Current.GetTargetMinutes(day.DayOfWeek);
                totalSoll += target;

                if (wd.ComeLocal.HasValue && wd.GoLocal.HasValue)
                {
                    var pauses = _workDays.GetBreaks(key).Where(b => b.EndLocal.HasValue).Sum(b => (int)(b.EndLocal!.Value - b.StartLocal).TotalMinutes);
                    totalNet += (int)(wd.GoLocal.Value - wd.ComeLocal.Value).TotalMinutes - pauses;
                }
            }

            NetMonthSummary = Fmt(totalNet);
            SollMonthSummary = Fmt(totalSoll);
            OvertimeMonthSummary = Fmt(totalNet - totalSoll);

            Breakdown.Clear();
            Breakdown.Add($"Ticket-Minuten: {TotalMinutes}");
            Breakdown.Add($"Netto Monat: {totalNet} min");
            Breakdown.Add($"Soll Monat: {totalSoll} min");
            Breakdown.Add($"Überstunden Monat: {totalNet - totalSoll} min");

            TopTasks.Clear();
            foreach (var (title, mins) in _tasks.GetTopTasksForMonth(month))
                TopTasks.Add(new ReportTaskItem { Title = title, Minutes = mins, DurationText = Fmt(mins) });
        }
        catch
        {
            MonthSummary = "0h 0m";
            NetMonthSummary = "0h 0m";
            SollMonthSummary = "0h 0m";
            OvertimeMonthSummary = "0h 0m";
            Breakdown.Clear();
            Breakdown.Add("Report konnte nicht geladen werden.");
            TopTasks.Clear();
        }
    }

    private static string Fmt(int mins) => $"{mins / 60}h {Math.Abs(mins % 60):00}m";
    public override string ToString() => Title;
}

public class ReportTaskItem
{
    public string Title { get; set; } = string.Empty;
    public int Minutes { get; set; }
    public string DurationText { get; set; } = string.Empty;
}
