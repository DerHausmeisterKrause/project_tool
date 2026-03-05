using System.Collections.Generic;
using TaskTool.Infrastructure;
using TaskTool.Services;

namespace TaskTool.ViewModels;

public class SettingsViewModel : ObservableObject
{
    private readonly SettingsService _settings;
    private readonly NotificationService _notifications;
    private readonly OutlookCalendarService _outlookCalendar;
    public string Title => "Einstellungen";

    public bool OutlookSyncEnabled { get => _settings.Current.OutlookSyncEnabled; set { _settings.Current.OutlookSyncEnabled = value; Save(); } }
    public string OutlookCategoryName { get => _settings.Current.OutlookCategoryName; set { _settings.Current.OutlookCategoryName = value; Save(); } }
    public bool OutlookCalendarEnabled { get => _settings.Current.OutlookCalendarEnabled; set { _settings.Current.OutlookCalendarEnabled = value; Save(); } }
    public bool OutlookConflictWarningsEnabled { get => _settings.Current.OutlookConflictWarningsEnabled; set { _settings.Current.OutlookConflictWarningsEnabled = value; Save(); } }
    public bool OutlookTeamsButtonEnabled { get => _settings.Current.OutlookTeamsButtonEnabled; set { _settings.Current.OutlookTeamsButtonEnabled = value; Save(); } }
    public string OutlookCalendarSyncMode { get => _settings.Current.OutlookCalendarSyncMode; set { _settings.Current.OutlookCalendarSyncMode = value; Save(); } }
    public int OutlookCalendarSyncIntervalMinutes { get => _settings.Current.OutlookCalendarSyncIntervalMinutes; set { _settings.Current.OutlookCalendarSyncIntervalMinutes = value; Save(); } }
    public int OutlookCalendarRangePastDays { get => _settings.Current.OutlookCalendarRangePastDays; set { _settings.Current.OutlookCalendarRangePastDays = value; Save(); } }
    public int OutlookCalendarRangeFutureDays { get => _settings.Current.OutlookCalendarRangeFutureDays; set { _settings.Current.OutlookCalendarRangeFutureDays = value; Save(); } }

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

    public SettingsViewModel(SettingsService settings, NotificationService notifications, OutlookCalendarService outlookCalendar)
    {
        _settings = settings;
        _notifications = notifications;
        _outlookCalendar = outlookCalendar;
        TestReminderCommand = new RelayCommand(() => _notifications.ShowTestNotification());
        RefreshOutlookCalendarCommand = new RelayCommand(async () => await _outlookCalendar.TriggerSyncAsync("manual-button"));
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
