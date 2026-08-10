using System.Collections.ObjectModel;
using TaskTool.Infrastructure;
using TaskTool.Services;

namespace TaskTool.ViewModels;

public class ReportsViewModel : ObservableObject
{
    private readonly TaskService _tasks;

    public string Title => "Reports";

    private int _ticketMinutesToday;
    public int TicketMinutesToday { get => _ticketMinutesToday; set => Set(ref _ticketMinutesToday, value); }

    private int _ticketMinutesCurrentMonth;
    public int TicketMinutesCurrentMonth { get => _ticketMinutesCurrentMonth; set => Set(ref _ticketMinutesCurrentMonth, value); }

    public ObservableCollection<ReportTaskItem> TopTasks { get; } = new();
    public RelayCommand RefreshCommand { get; }

    public ReportsViewModel(TaskService tasks)
    {
        _tasks = tasks;
        RefreshCommand = new RelayCommand(Refresh);
        Refresh();
    }

    private void Refresh()
    {
        try
        {
            var month = DateTime.Today;
            TicketMinutesToday = _tasks.GetTicketMinutesForDay(month);
            TicketMinutesCurrentMonth = _tasks.GetMonthTicketMinutes(month);

            TopTasks.Clear();
            foreach (var (title, mins) in _tasks.GetTopTasksForMonth(month))
                TopTasks.Add(new ReportTaskItem { Title = title, Minutes = mins, DurationText = $"{mins:N0} Min." });
        }
        catch
        {
            TicketMinutesToday = 0;
            TicketMinutesCurrentMonth = 0;
            TopTasks.Clear();
        }
    }

    public override string ToString() => Title;
}

public class ReportTaskItem
{
    public string Title { get; set; } = string.Empty;
    public int Minutes { get; set; }
    public string DurationText { get; set; } = string.Empty;
}
