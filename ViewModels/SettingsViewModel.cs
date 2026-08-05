using System;
using System.Collections.Generic;
using System.Windows;
using TaskTool.Infrastructure;
using TaskTool.Services;

namespace TaskTool.ViewModels;

public class SettingsViewModel : ObservableObject
{
    private readonly SettingsService _settings;
    private readonly NotificationService _notifications;
    private readonly OutlookCalendarService _outlookCalendar;
    private readonly TaskService _tasks;
    private readonly TicketSystemService _ticketSystem;
    public string Title => "Einstellungen";

    public bool OutlookSyncEnabled { get => _settings.Current.OutlookSyncEnabled; set { _settings.Current.OutlookSyncEnabled = value; Save(); } }
    public string OutlookCategoryName { get => _settings.Current.OutlookCategoryName; set { _settings.Current.OutlookCategoryName = value; Save(); } }
    public bool OutlookCalendarEnabled { get => _settings.Current.OutlookCalendarEnabled; set { _settings.Current.OutlookCalendarEnabled = value; Save(); } }
    public bool OutlookConflictWarningsEnabled { get => _settings.Current.OutlookConflictWarningsEnabled; set { _settings.Current.OutlookConflictWarningsEnabled = value; Save(); } }
    public bool OutlookTeamsButtonEnabled { get => _settings.Current.OutlookTeamsButtonEnabled; set { _settings.Current.OutlookTeamsButtonEnabled = value; Save(); } }
    public bool OutlookInterpretAllDayAsMarkers { get => _settings.Current.OutlookInterpretAllDayAsMarkers; set { _settings.Current.OutlookInterpretAllDayAsMarkers = value; Save(); } }
    public bool ShowWeekendInWeekView { get => _settings.Current.ShowWeekendInWeekView; set { _settings.Current.ShowWeekendInWeekView = value; Save(); } }
    public string OutlookCalendarSyncMode { get => _settings.Current.OutlookCalendarSyncMode; set { _settings.Current.OutlookCalendarSyncMode = value; Save(); } }
    public int OutlookCalendarSyncIntervalMinutes { get => _settings.Current.OutlookCalendarSyncIntervalMinutes; set { _settings.Current.OutlookCalendarSyncIntervalMinutes = value; Save(); } }
    public int OutlookCalendarRangePastDays { get => _settings.Current.OutlookCalendarRangePastDays; set { _settings.Current.OutlookCalendarRangePastDays = value; Save(); } }
    public int OutlookCalendarRangeFutureDays { get => _settings.Current.OutlookCalendarRangeFutureDays; set { _settings.Current.OutlookCalendarRangeFutureDays = value; Save(); } }

    public string TicketSystemWebUrl { get => _settings.Current.TicketSystemWebUrl; set { _settings.Current.TicketSystemWebUrl = value; Save(); } }
    public string TicketSystemApiUrl { get => _settings.Current.TicketSystemApiUrl; set { _settings.Current.TicketSystemApiUrl = value; Save(); } }
    public string TicketSystemUsername { get => _settings.Current.TicketSystemUsername; set { _settings.Current.TicketSystemUsername = value; Save(); } }
    public string TicketSystemPassword { get => _settings.Current.TicketSystemPassword; set { _settings.Current.TicketSystemPassword = value; Save(); } }
    public string TicketSystemApiToken { get => _settings.Current.TicketSystemApiToken; set { _settings.Current.TicketSystemApiToken = value; Save(); } }

    private string _ticketSystemStatus = string.Empty;
    public string TicketSystemStatus { get => _ticketSystemStatus; set => Set(ref _ticketSystemStatus, value); }

    public int ReminderLeadMinutes { get => _settings.Current.ReminderLeadMinutes; set { _settings.Current.ReminderLeadMinutes = value; Save(); } }
    public string DateTimeFormat { get => _settings.Current.DateTimeFormat; set { _settings.Current.DateTimeFormat = value; Save(); } }
    public int MondayTargetMinutes { get => _settings.Current.MondayTargetMinutes; set { _settings.Current.MondayTargetMinutes = value; Save(); } }
    public int TuesdayTargetMinutes { get => _settings.Current.TuesdayTargetMinutes; set { _settings.Current.TuesdayTargetMinutes = value; Save(); } }
    public int WednesdayTargetMinutes { get => _settings.Current.WednesdayTargetMinutes; set { _settings.Current.WednesdayTargetMinutes = value; Save(); } }
    public int ThursdayTargetMinutes { get => _settings.Current.ThursdayTargetMinutes; set { _settings.Current.ThursdayTargetMinutes = value; Save(); } }
    public int FridayTargetMinutes { get => _settings.Current.FridayTargetMinutes; set { _settings.Current.FridayTargetMinutes = value; Save(); } }
    public int SaturdayTargetMinutes { get => _settings.Current.SaturdayTargetMinutes; set { _settings.Current.SaturdayTargetMinutes = value; Save(); } }
    public int SundayTargetMinutes { get => _settings.Current.SundayTargetMinutes; set { _settings.Current.SundayTargetMinutes = value; Save(); } }
    public bool DynamicIslandEnabled { get => _settings.Current.DynamicIslandEnabled; set { _settings.Current.DynamicIslandEnabled = value; Save(); } }

    public List<string> OutlookSyncModes { get; } = new() { "Manual", "Periodic" };

    public RelayCommand TestReminderCommand { get; }
    public RelayCommand RefreshOutlookCalendarCommand { get; }
    public RelayCommand TestOutlookConnectionCommand { get; }
    public RelayCommand ImportTicketSystemTasksCommand { get; }

    public SettingsViewModel(SettingsService settings, NotificationService notifications, OutlookCalendarService outlookCalendar, TaskService tasks, TicketSystemService ticketSystem)
    {
        _settings = settings;
        _notifications = notifications;
        _outlookCalendar = outlookCalendar;
        _tasks = tasks;
        _ticketSystem = ticketSystem;
        TestReminderCommand = new RelayCommand(() => _notifications.ShowTestNotification());
        RefreshOutlookCalendarCommand = new RelayCommand(async () => await _outlookCalendar.TriggerSyncAsync("manual-button"));
        TestOutlookConnectionCommand = new RelayCommand(TestOutlookConnection);
        ImportTicketSystemTasksCommand = new RelayCommand(async () => await ImportTicketSystemTasksAsync());
    }

    private async Task ImportTicketSystemTasksAsync()
    {
        TicketSystemStatus = "Tickets werden abgerufen ...";
        var result = await _ticketSystem.ImportAssignedOpenTicketsAsync();
        TicketSystemStatus = string.IsNullOrWhiteSpace(_ticketSystem.LastError)
            ? $"Tickets importiert: {result.created} neu, {result.skipped} übersprungen."
            : _ticketSystem.LastError;
    }

    private void TestOutlookConnection()
    {
        var ok = _tasks.TestOutlookConnection();
        if (ok)
        {
            MessageBox.Show("Outlook Verbindungstest erfolgreich.", "Outlook Test", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var hex = ExtractHResultHex(_tasks.LastError);
        MessageBox.Show($"Outlook Test fehlgeschlagen. Details in logs.txt: {hex}", "Outlook Test", MessageBoxButton.OK, MessageBoxImage.Error);
    }

    private static string ExtractHResultHex(string error)
    {
        if (string.IsNullOrWhiteSpace(error)) return "0x00000000";
        var marker = "0x";
        var idx = error.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
        if (idx < 0) return "0x00000000";
        var end = idx + 2;
        while (end < error.Length && Uri.IsHexDigit(error[end])) end++;
        return error[idx..end];
    }

    private void Save()
    {
        _settings.Save();
        _notifications.HandleSettingsChanged();
        _outlookCalendar.HandleSettingsChanged();
        Raise(string.Empty);
    }

    public override string ToString() => Title;
}
