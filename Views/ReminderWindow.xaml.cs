using System.Windows;

namespace TaskTool.Views;

public partial class ReminderWindow : Window
{
    public event EventHandler? SnoozeRequested;

    public ReminderWindow(string title, DateTime start)
    {
        InitializeComponent();
        TaskLabel.Text = title;
        StartLabel.Text = $"Start: {start:HH:mm}";
    }

    private void Snooze_Click(object sender, RoutedEventArgs e)
    {
        SnoozeRequested?.Invoke(this, EventArgs.Empty);
        Close();
    }

    private void Close_Click(object sender, RoutedEventArgs e) => Close();
}
