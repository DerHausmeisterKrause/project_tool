using System;
using System.Collections.Generic;
using System.Windows;
using TaskTool.Infrastructure;
using TaskTool.Models;
using TaskTool.Services;

namespace TaskTool.ViewModels;

public class SettingsViewModel : ObservableObject
{
    private readonly SettingsService _settings;
    private readonly NotificationService _notifications;
    private readonly OutlookCalendarService _outlookCalendar;
    private readonly TaskService _tasks;
    private readonly TicketSystemService _ticketSystem;
    private readonly Action? _tasksChanged;
    private readonly UpdateService _updates;
    public string Title => "Einstellungen";
    public string InstalledVersion => _settings.Current.InstalledVersion;
    public bool CheckForUpdatesOnStartup { get => _settings.Current.CheckForUpdatesOnStartup; set { _settings.Current.CheckForUpdatesOnStartup = value; Save(); } }
    public bool AutoInstallUpdatesOnStartup { get => _settings.Current.AutoInstallUpdatesOnStartup; set { _settings.Current.AutoInstallUpdatesOnStartup = value; Save(); } }
    private UpdateState _updateState = UpdateState.Idle;
    public UpdateState CurrentUpdateState { get => _updateState; set { if (Set(ref _updateState, value)) RaiseUpdateCommands(); } }
    private string _updateStatus = "Noch nicht geprüft";
    public string UpdateStatus { get => _updateStatus; set => Set(ref _updateStatus, value); }
    private string _lastUpdateCheck = "Noch nicht geprüft";
    public string LastUpdateCheck { get => _lastUpdateCheck; set => Set(ref _lastUpdateCheck, value); }
    private string _availableVersion = "–";
    public string AvailableVersion { get => _availableVersion; set => Set(ref _availableVersion, value); }
    private string _releaseNotes = string.Empty;
    public string ReleaseNotes { get => _releaseNotes; set => Set(ref _releaseNotes, value); }
    private int _updateProgress;
    public int UpdateProgress { get => _updateProgress; set => Set(ref _updateProgress, value); }
    private UpdateInfo? _availableUpdate;

    public bool OutlookSyncEnabled { get => _settings.Current.OutlookSyncEnabled; set { _settings.Current.OutlookSyncEnabled = value; Save(); } }
    public string OutlookCategoryName { get => _settings.Current.OutlookCategoryName; set { _settings.Current.OutlookCategoryName = value; Save(); } }
    public bool OutlookCalendarEnabled { get => _settings.Current.OutlookCalendarEnabled; set { _settings.Current.OutlookCalendarEnabled = value; Save(); } }
    public bool OutlookConflictWarningsEnabled { get => _settings.Current.OutlookConflictWarningsEnabled; set { _settings.Current.OutlookConflictWarningsEnabled = value; Save(); } }
    public bool OutlookTeamsButtonEnabled { get => _settings.Current.OutlookTeamsButtonEnabled; set { _settings.Current.OutlookTeamsButtonEnabled = value; Save(); } }
    public bool OutlookInterpretAllDayAsMarkers { get => _settings.Current.OutlookInterpretAllDayAsMarkers; set { _settings.Current.OutlookInterpretAllDayAsMarkers = value; Save(); } }
    public bool ShowWeekendInWeekView { get => _settings.Current.ShowWeekendInWeekView; set { _settings.Current.ShowWeekendInWeekView = value; Save(); } }
    public bool ShowInternalTaskSegmentsInCalendar { get => _settings.Current.ShowInternalTaskSegmentsInCalendar; set { _settings.Current.ShowInternalTaskSegmentsInCalendar = value; Save(); } }
    public bool HidePastTodayItems { get => _settings.Current.HidePastTodayItems; set { _settings.Current.HidePastTodayItems = value; Save(); } }
    public string CalendarTimeZoneId { get => _settings.Current.CalendarTimeZoneId; set { _settings.Current.CalendarTimeZoneId = value; Save(); } }
    public string OutlookCalendarSyncMode { get => _settings.Current.OutlookCalendarSyncMode; set { _settings.Current.OutlookCalendarSyncMode = value; Save(); } }
    public int OutlookCalendarSyncIntervalMinutes { get => _settings.Current.OutlookCalendarSyncIntervalMinutes; set { _settings.Current.OutlookCalendarSyncIntervalMinutes = value; Save(); } }
    public int OutlookCalendarRangePastDays { get => _settings.Current.OutlookCalendarRangePastDays; set { _settings.Current.OutlookCalendarRangePastDays = value; Save(); } }
    public int OutlookCalendarRangeFutureDays { get => _settings.Current.OutlookCalendarRangeFutureDays; set { _settings.Current.OutlookCalendarRangeFutureDays = value; Save(); } }
    public int DefaultSegmentDurationMinutes { get => _settings.Current.DefaultSegmentDurationMinutes; set { _settings.Current.DefaultSegmentDurationMinutes = value; Save(); } }
    public string HomeOfficeMailRecipient1 { get => _settings.Current.HomeOfficeMailRecipient1; set { _settings.Current.HomeOfficeMailRecipient1 = value; Save(); } }
    public string HomeOfficeMailRecipient2 { get => _settings.Current.HomeOfficeMailRecipient2; set { _settings.Current.HomeOfficeMailRecipient2 = value; Save(); } }

    public string TicketSystemWebUrl { get => _settings.Current.TicketSystemWebUrl; set { _settings.Current.TicketSystemWebUrl = value; Save(); } }
    public string TicketSystemApiUrl { get => _settings.Current.TicketSystemApiUrl; set { _settings.Current.TicketSystemApiUrl = value; Save(); } }
    public string TicketSystemUsername { get => _settings.Current.TicketSystemUsername; set { _settings.Current.TicketSystemUsername = value; Save(); } }
    private const string TicketSystemPasswordMask = "••••••••";
    public string TicketSystemPassword
    {
        get => string.IsNullOrWhiteSpace(_settings.GetTicketSystemPassword()) ? string.Empty : TicketSystemPasswordMask;
        set
        {
            if (string.Equals(value, TicketSystemPasswordMask, StringComparison.Ordinal)) return;
            _settings.SetTicketSystemPassword(value ?? string.Empty);
            Save();
            Raise();
        }
    }
    public int TicketSystemAgentId { get => _settings.Current.TicketSystemAgentId; set { _settings.Current.TicketSystemAgentId = value; Save(); } }
    public string TicketSystemTicketSearchRoute { get => _settings.Current.TicketSystemTicketSearchRoute; set { _settings.Current.TicketSystemTicketSearchRoute = value; Save(); } }
    public string TicketSystemTicketSearchMethod { get => _settings.Current.TicketSystemTicketSearchMethod; set { _settings.Current.TicketSystemTicketSearchMethod = value; Save(); } }
    public string TicketSystemTicketSearchAuthMode { get => _settings.Current.TicketSystemTicketSearchAuthMode; set { _settings.Current.TicketSystemTicketSearchAuthMode = value; Save(); } }
    public string TicketSystemTicketGetRouteTemplate { get => _settings.Current.TicketSystemTicketGetRouteTemplate; set { _settings.Current.TicketSystemTicketGetRouteTemplate = value; Save(); } }
    public string TicketSystemTicketGetMethod { get => _settings.Current.TicketSystemTicketGetMethod; set { _settings.Current.TicketSystemTicketGetMethod = value; Save(); } }
    public string TicketSystemTicketGetAuthMode { get => _settings.Current.TicketSystemTicketGetAuthMode; set { _settings.Current.TicketSystemTicketGetAuthMode = value; Save(); } }
    public string TicketSystemTicketUpdateRoute { get => _settings.Current.TicketSystemTicketUpdateRoute; set { _settings.Current.TicketSystemTicketUpdateRoute = value; Save(); } }
    public string TicketSystemDynamicFieldOptionsRoute { get => _settings.Current.TicketSystemDynamicFieldOptionsRoute; set { _settings.Current.TicketSystemDynamicFieldOptionsRoute = value; Save(); } }
    public string TicketSystemCostCenterFieldName { get => _settings.Current.TicketSystemCostCenterFieldName; set { _settings.Current.TicketSystemCostCenterFieldName = value; Save(); } }
    public string TicketSystemOrderFieldName { get => _settings.Current.TicketSystemOrderFieldName; set { _settings.Current.TicketSystemOrderFieldName = value; Save(); } }
    public string TicketSystemCostCenterOptions { get => _settings.Current.TicketSystemCostCenterOptions; set { _settings.Current.TicketSystemCostCenterOptions = value; Save(); } }
    public string TicketSystemOrderOptions { get => _settings.Current.TicketSystemOrderOptions; set { _settings.Current.TicketSystemOrderOptions = value; Save(); } }
    public int TicketSystemSyncIntervalMinutes { get => _settings.Current.TicketSystemSyncIntervalMinutes; set { _settings.Current.TicketSystemSyncIntervalMinutes = value; Save(); } }
    public bool TicketSystemOnlyOpenTickets { get => _settings.Current.TicketSystemOnlyOpenTickets; set { _settings.Current.TicketSystemOnlyOpenTickets = value; Save(); } }
    public bool TicketSystemShowClosedTickets { get => _settings.Current.TicketSystemShowClosedTickets; set { _settings.Current.TicketSystemShowClosedTickets = value; Save(); } }
    public bool TicketSystemIncludeOwner { get => _settings.Current.TicketSystemIncludeOwner; set { _settings.Current.TicketSystemIncludeOwner = value; Save(); } }
    public bool TicketSystemIncludeResponsible { get => _settings.Current.TicketSystemIncludeResponsible; set { _settings.Current.TicketSystemIncludeResponsible = value; Save(); } }
    public bool NotifyOnNewAssignedTickets { get => _settings.Current.NotifyOnNewAssignedTickets; set { _settings.Current.NotifyOnNewAssignedTickets = value; Save(); } }
    public bool TicketSystemAutofillCredentials
    {
        get => _settings.Current.TicketSystemAutofillCredentials;
        set
        {
            _settings.Current.TicketSystemAutofillCredentials = value;
            if (!value) _settings.Current.TicketSystemAutoLogin = false;
            Save();
            Raise(nameof(TicketSystemAutoLogin));
        }
    }
    public bool TicketSystemAutoLogin
    {
        get => _settings.Current.TicketSystemAutoLogin;
        set
        {
            _settings.Current.TicketSystemAutoLogin = value;
            if (value) _settings.Current.TicketSystemAutofillCredentials = true;
            Save();
            Raise(nameof(TicketSystemAutofillCredentials));
        }
    }

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
    public List<string> CalendarTimeZones { get; } = new() { "Europe/Berlin", "Europe/London", "UTC", "Europe/Vienna", "Europe/Zurich" };
    public List<string> TicketSystemSearchMethods { get; } = new() { "POST", "GET" };
    public List<string> TicketSystemAuthModes { get; } = new() { "Session", "Direct" };
    public List<string> TicketSystemTicketGetAuthModes { get; } = new() { "Session", "Direct" };
    public List<int> SegmentDurationOptions { get; } = new() { 15, 30, 45, 60, 90, 120, 180, 240 };

    public RelayCommand TestReminderCommand { get; }
    public RelayCommand RefreshOutlookCalendarCommand { get; }
    public RelayCommand TestOutlookConnectionCommand { get; }
    public RelayCommand ImportTicketSystemTasksCommand { get; }
    public RelayCommand TestTicketSystemConnectionCommand { get; }
    public RelayCommand TestTicketSystemRoutesCommand { get; }
    public RelayCommand CheckForUpdatesCommand { get; }
    public RelayCommand InstallUpdateCommand { get; }
    public RelayCommand OpenReleaseCommand { get; }

    public SettingsViewModel(SettingsService settings, NotificationService notifications, OutlookCalendarService outlookCalendar, TaskService tasks, TicketSystemService ticketSystem, UpdateService updates, Action? tasksChanged = null)
    {
        _settings = settings;
        _notifications = notifications;
        _outlookCalendar = outlookCalendar;
        _tasks = tasks;
        _ticketSystem = ticketSystem;
        _updates = updates;
        _tasksChanged = tasksChanged;
        TestReminderCommand = new RelayCommand(() => _notifications.ShowTestNotification());
        RefreshOutlookCalendarCommand = new RelayCommand(async () => await _outlookCalendar.TriggerSyncAsync("manual-button"));
        TestOutlookConnectionCommand = new RelayCommand(TestOutlookConnection);
        ImportTicketSystemTasksCommand = new RelayCommand(async () => await ImportTicketSystemTasksAsync());
        TestTicketSystemConnectionCommand = new RelayCommand(async () => await TestTicketSystemConnectionAsync());
        TestTicketSystemRoutesCommand = new RelayCommand(async () => await TestTicketSystemRoutesAsync());
        CheckForUpdatesCommand = new RelayCommand(async () => await CheckForUpdatesAsync(false), () => CurrentUpdateState is not (UpdateState.Checking or UpdateState.Downloading or UpdateState.Installing));
        InstallUpdateCommand = new RelayCommand(async () => await InstallUpdateAsync(), () => CurrentUpdateState == UpdateState.UpdateAvailable && _availableUpdate != null);
        OpenReleaseCommand = new RelayCommand(() => { if (_availableUpdate != null) UrlLauncher.TryOpen(_availableUpdate.HtmlUrl, out _); }, () => _availableUpdate != null);
    }

    public async Task RunStartupUpdateCheckAsync()
    {
        if (!CheckForUpdatesOnStartup) return;
        await CheckForUpdatesAsync(true);
        if (CurrentUpdateState == UpdateState.Failed)
        {
            ServiceLocator.Logger.Error($"[StartupUpdate] failed='{UpdateStatus}'");
            return;
        }
        if (CurrentUpdateState != UpdateState.UpdateAvailable || _availableUpdate == null) return;
        ServiceLocator.Logger.Info($"[StartupUpdate] installedVersion={InstalledVersion} remoteVersion={_availableUpdate.Version} updateAvailable=true");
        ServiceLocator.Logger.Info($"[StartupUpdate] autoInstall={AutoInstallUpdatesOnStartup.ToString().ToLowerInvariant()}");
        if (!AutoInstallUpdatesOnStartup) return;
        await InstallUpdateAsync(true);
    }

    public void RefreshInstalledVersion() => Raise(nameof(InstalledVersion));

    private async Task CheckForUpdatesAsync(bool automatic)
    {
        CurrentUpdateState = UpdateState.Checking; UpdateStatus = "Updates werden geprüft ...";
        try
        {
            var result = await _updates.CheckForUpdatesAsync();
            LastUpdateCheck = DateTime.Now.ToString("dd.MM.yyyy HH:mm");
            _availableUpdate = result.Update;
            AvailableVersion = result.Update?.Version.ToString() ?? "–";
            ReleaseNotes = result.Update?.ReleaseNotes ?? string.Empty;
            CurrentUpdateState = result.UpdateAvailable ? UpdateState.UpdateAvailable : UpdateState.UpToDate;
            UpdateStatus = result.Message;
        }
        catch (Exception ex)
        {
            CurrentUpdateState = UpdateState.Failed; UpdateStatus = automatic ? "Automatische Prüfung fehlgeschlagen." : $"Updateprüfung fehlgeschlagen: {ex.Message}";
        }
        RaiseUpdateCommands();
    }

    private Task InstallUpdateAsync() => InstallUpdateAsync(false);

    private async Task InstallUpdateAsync(bool startupAutomatic)
    {
        if (_availableUpdate == null) return;
        try
        {
            CurrentUpdateState = UpdateState.Downloading; UpdateStatus = "Update wird heruntergeladen ...";
            if (startupAutomatic) ServiceLocator.Logger.Info("[StartupUpdate] downloadStarted=true");
            var progress = new Progress<int>(value => { UpdateProgress = value; UpdateStatus = $"Update wird heruntergeladen ... {value} %"; });
            var path = await _updates.DownloadUpdateAsync(_availableUpdate, progress);
            CurrentUpdateState = UpdateState.ReadyToInstall;
            if (!_updates.InstallUpdate(path, _availableUpdate, out var error))
            {
                CurrentUpdateState = UpdateState.Failed;
                UpdateStatus = error;
                if (startupAutomatic) ServiceLocator.Logger.Error($"[StartupUpdate] failed='{error}'");
                return;
            }
            if (startupAutomatic) ServiceLocator.Logger.Info("[StartupUpdate] installPrepared=true");
            CurrentUpdateState = UpdateState.Installing; UpdateStatus = "Update wird installiert ...";
            Application.Current.Shutdown();
        }
        catch (Exception ex)
        {
            CurrentUpdateState = UpdateState.Failed;
            UpdateStatus = $"Update fehlgeschlagen: {ex.Message}";
            if (startupAutomatic) ServiceLocator.Logger.Error($"[StartupUpdate] failed='{ex.Message}'");
        }
        RaiseUpdateCommands();
    }

    private void RaiseUpdateCommands()
    {
        CheckForUpdatesCommand?.RaiseCanExecuteChanged(); InstallUpdateCommand?.RaiseCanExecuteChanged(); OpenReleaseCommand?.RaiseCanExecuteChanged();
    }

    private async Task TestTicketSystemRoutesAsync()
    {
        TicketSystemStatus = "Znuny API-Routen werden getestet ...";
        var result = await _ticketSystem.TestRoutesAsync();
        TicketSystemStatus = result.message;
        MessageBox.Show(result.message, "Znuny API-Routentest", MessageBoxButton.OK, result.success ? MessageBoxImage.Information : MessageBoxImage.Warning);
    }

    private async Task TestTicketSystemConnectionAsync()
    {
        TicketSystemStatus = "Znuny-Verbindung wird getestet ...";
        var result = await _ticketSystem.TestConnectionAsync();
        TicketSystemStatus = result.message;
        MessageBox.Show(result.message, "Znuny Verbindungstest", MessageBoxButton.OK, result.success ? MessageBoxImage.Information : MessageBoxImage.Warning);
    }

    private async Task ImportTicketSystemTasksAsync()
    {
        TicketSystemStatus = "Tickets werden abgerufen ...";
        var result = await _ticketSystem.ImportAssignedOpenTicketsAsync();
        if (string.IsNullOrWhiteSpace(_ticketSystem.LastError))
        {
            _tasksChanged?.Invoke();
            TicketSystemStatus = $"Znuny Sync fertig: {result.created} neu, {result.updated} aktualisiert, {result.skipped} übersprungen.";
            return;
        }

        TicketSystemStatus = _ticketSystem.LastError;
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
        _ticketSystem.HandleSettingsChanged();
        Raise(string.Empty);
    }

    public override string ToString() => Title;
}
