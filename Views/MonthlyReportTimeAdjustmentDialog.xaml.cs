using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using TaskTool.Services;

namespace TaskTool.Views;

public partial class MonthlyReportTimeAdjustmentDialog : Window
{
    private readonly TaskService _tasks;
    private readonly InputModel _model;

    public MonthlyReportTimeAdjustmentDialog(TaskService tasks, DateTime initialMonth)
    {
        InitializeComponent();
        _tasks = tasks;
        _model = new InputModel
        {
            Month = new DateTime(initialMonth.Year, initialMonth.Month, 1),
            Today = DateTime.Today
        };
        DataContext = _model;
    }

    private void Add_Click(object sender, RoutedEventArgs e)
    {
        if (_model.Month is not DateTime selected)
        {
            _model.Error = "Bitte einen Monat auswählen.";
            return;
        }
        var month = new DateTime(selected.Year, selected.Month, 1);
        var currentMonth = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
        if (month > currentMonth)
        {
            _model.Error = "Zukünftige Monate sind nicht zulässig.";
            return;
        }
        if (!TryCalculateSeconds(_model.Hours, _model.Minutes, out var seconds, out var validationError))
        {
            _model.Error = validationError;
            return;
        }

        try
        {
            _tasks.AddMonthlyReportTimeAdjustment(month, seconds, _model.Note);
            DialogResult = true;
        }
        catch (Exception)
        {
            _model.Error = "Die Nachbuchung konnte nicht gespeichert werden.";
        }
    }

    internal static bool TryCalculateSeconds(string hoursText, string minutesText, out long seconds, out string error)
    {
        seconds = 0;
        if (!int.TryParse(hoursText, out var hours) || hours < 0
            || !int.TryParse(minutesText, out var minutes) || minutes is < 0 or > 59)
        {
            error = "Stunden müssen mindestens 0 und Minuten zwischen 0 und 59 liegen.";
            return false;
        }
        seconds = checked((hours * 3600L) + (minutes * 60L));
        error = seconds > 0 ? string.Empty : "Die Nachbuchung muss größer als 0 Minuten sein.";
        return seconds > 0;
    }

    private sealed class InputModel : INotifyPropertyChanged
    {
        private string _error = string.Empty;
        public DateTime? Month { get; set; }
        public DateTime Today { get; set; }
        public string Hours { get; set; } = "0";
        public string Minutes { get; set; } = "0";
        public string Note { get; set; } = string.Empty;
        public string Error { get => _error; set { _error = value; PropertyChanged?.Invoke(this, new(nameof(Error))); } }
        public event PropertyChangedEventHandler? PropertyChanged;
    }
}
