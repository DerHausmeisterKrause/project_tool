using System.Collections.ObjectModel;
using System.Globalization;
using System.Windows;
using TaskTool.Models;
using TaskTool.Services;

namespace TaskTool.Views;

public partial class MonthlyReportTimeAdjustmentsDialog : Window
{
    private readonly TaskService _tasks;
    private readonly DateTime _month;
    private readonly DialogModel _model;
    public bool Changed { get; private set; }

    public MonthlyReportTimeAdjustmentsDialog(TaskService tasks, DateTime month)
    {
        InitializeComponent();
        _tasks = tasks;
        _month = new DateTime(month.Year, month.Month, 1);
        _model = new DialogModel { Heading = $"Nachbuchungen · {_month.ToString("MMMM yyyy", CultureInfo.GetCultureInfo("de-DE"))}" };
        DataContext = _model;
        Reload();
    }

    private void Delete_Click(object sender, RoutedEventArgs e)
    {
        if ((sender as FrameworkElement)?.Tag is not MonthlyReportTimeAdjustment adjustment) return;
        if (MessageBox.Show("Diese manuelle Nachbuchung wirklich löschen?", "Nachbuchung löschen", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes) return;
        _tasks.DeleteMonthlyReportTimeAdjustment(adjustment.Id);
        Changed = true;
        Reload();
    }

    private void Reload()
    {
        _model.Adjustments.Clear();
        foreach (var adjustment in _tasks.GetMonthlyReportTimeAdjustments(_month))
            _model.Adjustments.Add(adjustment);
        _model.EmptyVisibility = _model.Adjustments.Count == 0 ? Visibility.Visible : Visibility.Collapsed;
    }

    private sealed class DialogModel : TaskTool.Infrastructure.ObservableObject
    {
        private Visibility _emptyVisibility;
        public string Heading { get; init; } = string.Empty;
        public ObservableCollection<MonthlyReportTimeAdjustment> Adjustments { get; } = new();
        public Visibility EmptyVisibility { get => _emptyVisibility; set => Set(ref _emptyVisibility, value); }
    }
}
