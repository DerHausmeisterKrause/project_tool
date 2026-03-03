using TaskTool.Infrastructure;
using TaskTool.Services;

namespace TaskTool.ViewModels;

public class SettingsViewModel : ObservableObject
{
    private readonly SettingsService _settings;
    public string Title => "Einstellungen";

    public bool OutlookSyncEnabled { get => _settings.Current.OutlookSyncEnabled; set { _settings.Current.OutlookSyncEnabled = value; Save(); } }
    public string OutlookCategoryName { get => _settings.Current.OutlookCategoryName; set { _settings.Current.OutlookCategoryName = value; Save(); } }
    public int ReminderLeadMinutes { get => _settings.Current.ReminderLeadMinutes; set { _settings.Current.ReminderLeadMinutes = value; Save(); } }
    public string DateTimeFormat { get => _settings.Current.DateTimeFormat; set { _settings.Current.DateTimeFormat = value; Save(); } }
    public int MondayTargetMinutes { get => _settings.Current.MondayTargetMinutes; set { _settings.Current.MondayTargetMinutes = value; Save(); } }
    public int TuesdayTargetMinutes { get => _settings.Current.TuesdayTargetMinutes; set { _settings.Current.TuesdayTargetMinutes = value; Save(); } }
    public int WednesdayTargetMinutes { get => _settings.Current.WednesdayTargetMinutes; set { _settings.Current.WednesdayTargetMinutes = value; Save(); } }
    public int ThursdayTargetMinutes { get => _settings.Current.ThursdayTargetMinutes; set { _settings.Current.ThursdayTargetMinutes = value; Save(); } }
    public int FridayTargetMinutes { get => _settings.Current.FridayTargetMinutes; set { _settings.Current.FridayTargetMinutes = value; Save(); } }
    public int SaturdayTargetMinutes { get => _settings.Current.SaturdayTargetMinutes; set { _settings.Current.SaturdayTargetMinutes = value; Save(); } }
    public int SundayTargetMinutes { get => _settings.Current.SundayTargetMinutes; set { _settings.Current.SundayTargetMinutes = value; Save(); } }

    public SettingsViewModel(SettingsService settings) => _settings = settings;

    private void Save()
    {
        _settings.Save();
        Raise(string.Empty);
    }

    public override string ToString() => Title;
}
