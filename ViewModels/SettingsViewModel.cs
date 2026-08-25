using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Text.Json;
using System.Windows;
using System.Windows.Threading;
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
    private readonly DispatcherTimer _hourlyUpdateTimer;
    private const string ShortcutPasswordMask = "••••••••";
    public ObservableCollection<WebShortcutEditorViewModel> WebShortcuts { get; }
    private WebShortcutEditorViewModel? _selectedWebShortcut; public WebShortcutEditorViewModel? SelectedWebShortcut { get=>_selectedWebShortcut; set=>Set(ref _selectedWebShortcut,value); }
    private string _webShortcutStatus=""; public string WebShortcutStatus { get=>_webShortcutStatus; set=>Set(ref _webShortcutStatus,value); }
    public RelayCommand AddWebShortcutCommand { get; } public RelayCommand SaveWebShortcutCommand { get; } public RelayCommand RemoveWebShortcutCommand { get; }
    public ObservableCollection<WikiSourceEditorViewModel> WikiSources { get; }
    private WikiSourceEditorViewModel? _selectedWikiSource;
    public WikiSourceEditorViewModel? SelectedWikiSource { get => _selectedWikiSource; set { if (Set(ref _selectedWikiSource, value)) Raise(nameof(WikiIndexStatus)); } }
    private const string WikiSecretMask = "••••••••";
    public List<WikiChoice> WikiProviderTypes { get; } = new() { new("ConfluenceDataCenter", "Confluence Data Center"), new("ConfluenceCloud", "Confluence Cloud"), new("GenericRest", "Generic REST"), new("XWiki", "XWiki") };
    public List<WikiChoice> WikiAuthModes { get; } = new() { new("BearerToken", "Bearer Token"), new("UsernameToken", "Username + Token / Passwort"), new("Basic", "Basic Auth"), new("ApiKey", "API-Key Header"), new("WindowsIntegrated", "Windows Integrated") };
    public List<WikiChoice> WikiBrowserLoginModes { get; } = new() { new("BrowserSession", "Browser Session / Cookies"), new("WindowsIntegrated", "Windows Integrated"), new("UsernamePassword", "Benutzername + Passwort"), new("None", "Keine automatische Anmeldung") };
    public string CurrentWindowsUser => $"{Environment.UserDomainName}\\{Environment.UserName}";
    private string _wikiSettingsStatus = string.Empty;
    public string WikiSettingsStatus { get => _wikiSettingsStatus; set => Set(ref _wikiSettingsStatus, value); }
    private string _wikiTestSearchTerm = "Linux";
    public string WikiTestSearchTerm { get => _wikiTestSearchTerm; set => Set(ref _wikiTestSearchTerm, value); }
    private bool _isWikiTestRunning;
    public RelayCommand AddWikiSourceCommand { get; }
    public RelayCommand RemoveWikiSourceCommand { get; }
    public RelayCommand SaveWikiSourceCommand { get; }
    public RelayCommand DiscardWikiSourceCommand { get; }
    public RelayCommand TestWikiConnectionCommand { get; }
    public RelayCommand TestWikiSearchCommand { get; }
    public RelayCommand RefreshWikiIndexCommand { get; }
    public string WikiIndexStatus => SelectedWikiSource == null ? "Kein Wiki ausgewählt." : FormatWikiIndexStatus(SelectedWikiSource.ToModel());
    private readonly SemaphoreSlim _updateCheckGate = new(1, 1);
    public string Title => "Einstellungen";
    public string InstalledVersion => ServiceLocator.AppVersion.InstalledVersionText;
    public string UpdateChannel { get => _settings.Current.UpdateChannel; set { var normalized = value == "Pre-Release / Tester" || value == "PreRelease" ? "PreRelease" : "Stable"; if (_settings.Current.UpdateChannel == normalized) return; _settings.Current.UpdateChannel = normalized; Save(); Raise(); Raise(nameof(IsPreReleaseChannel)); _ = CheckForUpdatesAsync(false); } }
    public bool IsPreReleaseChannel => UpdateChannel == "PreRelease";
    public IReadOnlyList<WikiChoice> UpdateChannels { get; } = new[] { new WikiChoice("Stable", "Stable"), new WikiChoice("PreRelease", "Pre-Release / Tester") };
    public bool CheckForUpdatesOnStartup { get => _settings.Current.CheckForUpdatesOnStartup; set { _settings.Current.CheckForUpdatesOnStartup = value; Save(); } }
    public bool AutoInstallUpdatesOnStartup { get => _settings.Current.AutoInstallUpdatesOnStartup; set { _settings.Current.AutoInstallUpdatesOnStartup = value; Save(); } }
    public string LogLevel { get => _settings.Current.LogLevel; set { _settings.Current.LogLevel = value; Save(); } }
    public bool NotificationSoundEnabled { get => _settings.Current.NotificationSoundEnabled; set { _settings.Current.NotificationSoundEnabled = value; Save(); } }
    private UpdateState _updateState = UpdateState.Idle;
    public UpdateState CurrentUpdateState { get => _updateState; set { if (Set(ref _updateState, value)) { RaiseUpdateCommands(); RaiseUpdateBanner(); } } }
    private string _updateStatus = "Noch nicht geprüft";
    public string UpdateStatus { get => _updateStatus; set => Set(ref _updateStatus, value); }
    private string _lastUpdateCheck = "Noch nicht geprüft";
    public string LastUpdateCheck { get => _lastUpdateCheck; set => Set(ref _lastUpdateCheck, value); }
    private string _availableVersion = "–";
    public string AvailableVersion { get => _availableVersion; set { if (Set(ref _availableVersion, value)) RaiseUpdateBanner(); } }
    private string _releaseNotes = string.Empty;
    public string ReleaseNotes { get => _releaseNotes; set => Set(ref _releaseNotes, value); }
    private int _updateProgress;
    public int UpdateProgress { get => _updateProgress; set => Set(ref _updateProgress, value); }
    private UpdateInfo? _availableUpdate;
    public bool IsUpdateAvailable => CurrentUpdateState == UpdateState.UpdateAvailable && _availableUpdate != null;
    public string UpdateBannerText => IsUpdateAvailable ? $"Update {AvailableVersion} verfügbar" : string.Empty;

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
    public int TicketSystemCandidateUserId { get => _settings.Current.TicketSystemCandidateUserId; set { _settings.Current.TicketSystemCandidateUserId = value; Save(); } }
    public string TicketSystemCandidateKeywords { get => _settings.Current.TicketSystemCandidateKeywords; set { _settings.Current.TicketSystemCandidateKeywords = value; Save(); } }
    public string TicketSystemCandidateExcludeKeywords { get => _settings.Current.TicketSystemCandidateExcludeKeywords; set { _settings.Current.TicketSystemCandidateExcludeKeywords = value ?? string.Empty; Save(); } }
    public string TicketSystemTicketSearchRoute { get => _settings.Current.TicketSystemTicketSearchRoute; set { _settings.Current.TicketSystemTicketSearchRoute = value; Save(); } }
    public string TicketSystemTicketSearchMethod { get => _settings.Current.TicketSystemTicketSearchMethod; set { _settings.Current.TicketSystemTicketSearchMethod = value; Save(); } }
    public string TicketSystemTicketSearchAuthMode { get => _settings.Current.TicketSystemTicketSearchAuthMode; set { _settings.Current.TicketSystemTicketSearchAuthMode = value; Save(); } }
    public string TicketSystemTicketGetRouteTemplate { get => _settings.Current.TicketSystemTicketGetRouteTemplate; set { _settings.Current.TicketSystemTicketGetRouteTemplate = value; Save(); } }
    public string TicketSystemTicketGetMethod { get => _settings.Current.TicketSystemTicketGetMethod; set { _settings.Current.TicketSystemTicketGetMethod = value; Save(); } }
    public string TicketSystemTicketGetAuthMode { get => _settings.Current.TicketSystemTicketGetAuthMode; set { _settings.Current.TicketSystemTicketGetAuthMode = value; Save(); } }
    public string TicketSystemTicketUpdateRoute { get => _settings.Current.TicketSystemTicketUpdateRoute; set { _settings.Current.TicketSystemTicketUpdateRoute = value; Save(); } }
    public string TicketSystemTicketCreateRoute { get => _settings.Current.TicketSystemTicketCreateRoute; set { _settings.Current.TicketSystemTicketCreateRoute = value; Save(); } }
    public string TicketSystemTicketCreateMethod { get => _settings.Current.TicketSystemTicketCreateMethod; set { _settings.Current.TicketSystemTicketCreateMethod = value; Save(); } }
    public string TicketSystemCreateQueue { get => _settings.Current.TicketSystemCreateQueue; set { _settings.Current.TicketSystemCreateQueue = value; Save(); } }
    public string TicketSystemCreateState { get => _settings.Current.TicketSystemCreateState; set { _settings.Current.TicketSystemCreateState = value; Save(); } }
    public string TicketSystemCreatePriority { get => _settings.Current.TicketSystemCreatePriority; set { _settings.Current.TicketSystemCreatePriority = value; Save(); } }
    public string TicketSystemCreateType { get => _settings.Current.TicketSystemCreateType; set { _settings.Current.TicketSystemCreateType = value; Save(); } }
    public string TicketSystemCreateCustomerUser { get => _settings.Current.TicketSystemCreateCustomerUser; set { _settings.Current.TicketSystemCreateCustomerUser = value; Save(); } }
    public string TicketSystemReplyTemplate { get => _settings.Current.TicketSystemReplyTemplate; set { _settings.Current.TicketSystemReplyTemplate = value ?? string.Empty; Save(); } }
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
    public List<string> LogLevels { get; } = new() { "Info", "Warning", "Error" };
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
        WebShortcuts = new(_settings.Current.WebShortcuts.Select(x=>WebShortcutEditorViewModel.From(x,ShortcutPasswordMask))); SelectedWebShortcut=WebShortcuts.FirstOrDefault();
        WikiSources = new ObservableCollection<WikiSourceEditorViewModel>(_settings.Current.WikiSources.Select(x => WikiSourceEditorViewModel.FromModel(x, WikiSecretMask, x.Id == _settings.Current.DefaultWikiSourceId)));
        SelectedWikiSource = WikiSources.FirstOrDefault();
        _hourlyUpdateTimer = new DispatcherTimer { Interval = TimeSpan.FromHours(1) };
        _hourlyUpdateTimer.Tick += async (_, _) => await RunHourlyUpdateCheckAsync();
        TestReminderCommand = new RelayCommand(() => _notifications.ShowTestNotification());
        RefreshOutlookCalendarCommand = new RelayCommand(async () => await _outlookCalendar.TriggerSyncAsync("manual-button"));
        TestOutlookConnectionCommand = new RelayCommand(TestOutlookConnection);
        ImportTicketSystemTasksCommand = new RelayCommand(async () => await ImportTicketSystemTasksAsync());
        TestTicketSystemConnectionCommand = new RelayCommand(async () => await TestTicketSystemConnectionAsync());
        TestTicketSystemRoutesCommand = new RelayCommand(async () => await TestTicketSystemRoutesAsync());
        CheckForUpdatesCommand = new RelayCommand(async () => await CheckForUpdatesAsync(false), () => CurrentUpdateState is not (UpdateState.Checking or UpdateState.Downloading or UpdateState.Installing));
        InstallUpdateCommand = new RelayCommand(async () => await InstallUpdateAsync(), () => CurrentUpdateState == UpdateState.UpdateAvailable && _availableUpdate != null);
        OpenReleaseCommand = new RelayCommand(() => { if (_availableUpdate != null) UrlLauncher.TryOpen(_availableUpdate.HtmlUrl, out _); }, () => _availableUpdate != null);
        AddWikiSourceCommand = new RelayCommand(AddWikiSource);
        RemoveWikiSourceCommand = new RelayCommand(RemoveWikiSource);
        SaveWikiSourceCommand = new RelayCommand(SaveWikiSource);
        DiscardWikiSourceCommand = new RelayCommand(DiscardWikiSource);
        TestWikiConnectionCommand = new RelayCommand(async () => await TestWikiAsync(false), () => !_isWikiTestRunning);
        TestWikiSearchCommand = new RelayCommand(async () => await TestWikiAsync(true), () => !_isWikiTestRunning);
        RefreshWikiIndexCommand = new RelayCommand(async () => await RefreshWikiIndexAsync());
        AddWebShortcutCommand=new RelayCommand(()=>{var x=new WebShortcutEditorViewModel();WebShortcuts.Add(x);SelectedWebShortcut=x;WebShortcutStatus="Webseite angelegt. Bitte speichern.";}); SaveWebShortcutCommand=new RelayCommand(SaveWebShortcut); RemoveWebShortcutCommand=new RelayCommand(RemoveWebShortcut);
    }

    public void StartHourlyUpdateMonitor()
    {
        if (!_hourlyUpdateTimer.IsEnabled)
            _hourlyUpdateTimer.Start();
    }

    public void StopHourlyUpdateMonitor() => _hourlyUpdateTimer.Stop();

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

    private async Task CheckForUpdatesAsync(bool automatic, bool background = false)
    {
        if (!await _updateCheckGate.WaitAsync(0))
            return;

        var previousState = CurrentUpdateState;
        var previousStatus = UpdateStatus;
        try
        {
            if (!background)
            {
                CurrentUpdateState = UpdateState.Checking;
                UpdateStatus = "Updates werden geprüft ...";
            }

            var result = await _updates.CheckForUpdatesAsync();
            LastUpdateCheck = DateTime.Now.ToString("dd.MM.yyyy HH:mm");
            _availableUpdate = result.Update;
            AvailableVersion = result.Update?.Version.ToString() ?? "–";
            ReleaseNotes = result.Update?.ReleaseNotes ?? string.Empty;
            CurrentUpdateState = result.UpdateAvailable ? UpdateState.UpdateAvailable : UpdateState.UpToDate;
            UpdateStatus = result.Message;
            if (background)
            {
                ServiceLocator.Logger.Info($"[UpdateMonitor] action=hourly-check installedVersion={InstalledVersion} remoteVersion={AvailableVersion} updateAvailable={result.UpdateAvailable.ToString().ToLowerInvariant()}");
            }
        }
        catch (Exception ex)
        {
            if (background)
            {
                CurrentUpdateState = previousState;
                UpdateStatus = previousStatus;
                ServiceLocator.Logger.Warning($"[UpdateMonitor] hourlyCheckFailed message='{ex.Message}'");
            }
            else
            {
                CurrentUpdateState = UpdateState.Failed;
                UpdateStatus = automatic ? "Automatische Prüfung fehlgeschlagen." : $"Updateprüfung fehlgeschlagen: {ex.Message}";
            }
        }
        finally
        {
            _updateCheckGate.Release();
            RaiseUpdateCommands();
            RaiseUpdateBanner();
        }
    }

    private async Task RunHourlyUpdateCheckAsync()
    {
        if (CurrentUpdateState is UpdateState.Checking or UpdateState.Downloading or UpdateState.ReadyToInstall or UpdateState.Installing)
            return;
        await CheckForUpdatesAsync(automatic: true, background: true);
    }

    private void RaiseUpdateBanner()
    {
        Raise(nameof(IsUpdateAvailable));
        Raise(nameof(UpdateBannerText));
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

    private void AddWikiSource()
    {
        var editor = new WikiSourceEditorViewModel();
        WikiSources.Add(editor);
        SelectedWikiSource = editor;
        WikiSettingsStatus = "Neue Wiki-Quelle angelegt. Bitte Konfiguration ausfüllen und speichern.";
    }

    private void RemoveWikiSource()
    {
        var editor = SelectedWikiSource;
        if (editor == null) { WikiSettingsStatus = "Bitte zuerst eine Wiki-Quelle auswählen."; return; }
        if (MessageBox.Show($"Wiki '{editor.Name}' wirklich entfernen?", "Wiki entfernen", MessageBoxButton.YesNo, MessageBoxImage.Warning) != MessageBoxResult.Yes) return;
        var previous = _settings.Current.WikiSources.ToList();
        _settings.Current.WikiSources.RemoveAll(x => string.Equals(x.Id, editor.Id, StringComparison.Ordinal));
        if (_settings.Current.DefaultWikiSourceId == editor.Id) _settings.Current.DefaultWikiSourceId = _settings.Current.WikiSources.FirstOrDefault(x => x.Enabled)?.Id ?? string.Empty;
        if (!_settings.TrySave()) { _settings.Current.WikiSources = previous; WikiSettingsStatus = "Wiki konnte nicht entfernt werden. Bitte logs.txt prüfen."; return; }
        WikiSources.Remove(editor); SelectedWikiSource = WikiSources.FirstOrDefault();
        NotifySettingsConsumers(); WikiSettingsStatus = $"Wiki '{editor.Name}' wurde entfernt.";
    }

    private void DiscardWikiSource()
    {
        var editor = SelectedWikiSource; if (editor == null) return;
        var saved = _settings.Current.WikiSources.FirstOrDefault(x => x.Id == editor.Id);
        var replacement = saved == null ? new WikiSourceEditorViewModel() : WikiSourceEditorViewModel.FromModel(saved, WikiSecretMask, saved.Id == _settings.Current.DefaultWikiSourceId);
        var index = WikiSources.IndexOf(editor); WikiSources[index] = replacement; SelectedWikiSource = replacement;
        WikiSettingsStatus = "Nicht gespeicherte Änderungen wurden verworfen.";
    }

    private void SaveWikiSource()
    {
        if (!TryCreateWikiSourceFromEditor(out var source, out var error)) { WikiSettingsStatus = error; return; }
        var previous = _settings.Current.WikiSources.ToList();
        var previousSource = previous.FirstOrDefault(x => x.Id == source.Id);
        var searchConfigurationChanged = previousSource != null && WikiSearchConfigurationFingerprint(previousSource) != WikiSearchConfigurationFingerprint(source);
        var apiAccessChanged = previousSource != null && string.Join("|", previousSource.AuthMode, previousSource.Username, previousSource.SecretEncrypted) != string.Join("|", source.AuthMode, source.Username, source.SecretEncrypted);
        var index = _settings.Current.WikiSources.FindIndex(x => x.Id == source.Id);
        if (index >= 0) _settings.Current.WikiSources[index] = source; else _settings.Current.WikiSources.Add(source);
        if (SelectedWikiSource!.IsDefault || _settings.Current.WikiSources.Count(x => x.Enabled) == 1) { _settings.Current.DefaultWikiSourceId = source.Id; foreach (var editor in WikiSources) editor.IsDefault = editor.Id == source.Id; }
        if (!_settings.TrySave()) { _settings.Current.WikiSources = previous; WikiSettingsStatus = "Wiki konnte nicht gespeichert werden. Bitte logs.txt prüfen."; return; }
        SelectedWikiSource!.SetEncryptedSecret(source.SecretEncrypted); SelectedWikiSource.Secret = string.IsNullOrWhiteSpace(source.SecretEncrypted) ? string.Empty : WikiSecretMask;
        SelectedWikiSource.SetEncryptedBrowserPassword(source.BrowserPasswordEncrypted); SelectedWikiSource.BrowserPassword = string.IsNullOrWhiteSpace(source.BrowserPasswordEncrypted) ? string.Empty : WikiSecretMask;
        if (searchConfigurationChanged) ServiceLocator.WikiSearch.InvalidateSource(source.Id);
        else if (apiAccessChanged) ServiceLocator.WikiSearch.ResetFailedRunsForSource(source.Id);
        NotifySettingsConsumers(); WikiSettingsStatus = $"Wiki '{source.Name}' wurde gespeichert.";
        if (searchConfigurationChanged || previousSource == null) { ServiceLocator.WikiVocabulary.Invalidate(source.Id); _ = ServiceLocator.WikiVocabulary.RefreshAsync(source); }
        Raise(nameof(WikiIndexStatus));
    }

    private static string WikiSearchConfigurationFingerprint(WikiSourceSettings source)
        => string.Join("|", WikiScopePolicy.Fingerprint(source), source.HttpMethod, source.SearchUrlTemplate);

    private bool TryCreateWikiSourceFromEditor(out WikiSourceSettings source, out string error)
    {
        if (SelectedWikiSource == null) { source = new WikiSourceSettings(); error = "Bitte zuerst eine Wiki-Quelle auswählen."; return false; }
        source = SelectedWikiSource.ToModel();
        var secret = SelectedWikiSource.Secret;
        if (!string.Equals(secret, WikiSecretMask, StringComparison.Ordinal)) _settings.SetWikiSecret(source, secret ?? string.Empty);
        var browserPassword = SelectedWikiSource.BrowserPassword;
        if (!string.Equals(browserPassword, WikiSecretMask, StringComparison.Ordinal)) _settings.SetWikiBrowserPassword(source, browserPassword ?? string.Empty);
        return WikiSourceValidation.TryValidate(source, out error);
    }

    private async Task TestWikiAsync(bool includeExamples)
    {
        if (!TryCreateWikiSourceFromEditor(out var source, out var error)) { WikiSettingsStatus = error; return; }
        if (string.IsNullOrWhiteSpace(WikiTestSearchTerm)) { WikiSettingsStatus = "Bitte einen Test-Suchbegriff eingeben."; return; }
        _isWikiTestRunning = true; TestWikiConnectionCommand.RaiseCanExecuteChanged(); TestWikiSearchCommand.RaiseCanExecuteChanged();
        WikiSettingsStatus = includeExamples ? "Testsuche läuft …" : "Verbindung wird getestet …"; var watch = Stopwatch.StartNew();
        try
        {
            var results = await ServiceLocator.WikiSearch.TestSourceAsync(source, WikiTestSearchTerm);
            var examples = includeExamples ? string.Join(Environment.NewLine, results.Take(3).Select(x => $"• {x.Title}")) : string.Empty;
            WikiSettingsStatus = $"Verbindung erfolgreich · {results.Count} Treffer · {watch.ElapsedMilliseconds} ms" + (string.IsNullOrWhiteSpace(examples) ? string.Empty : Environment.NewLine + examples);
        }
        catch (Exception ex) { WikiSettingsStatus = DescribeWikiTestError(ex); }
        finally { _isWikiTestRunning = false; TestWikiConnectionCommand.RaiseCanExecuteChanged(); TestWikiSearchCommand.RaiseCanExecuteChanged(); }
    }

    private async Task RefreshWikiIndexAsync()
    {
        if (!TryCreateWikiSourceFromEditor(out var source, out var error)) { WikiSettingsStatus = error; return; }
        WikiSettingsStatus = "Wiki-Suchindex wird aktualisiert …"; await ServiceLocator.WikiVocabulary.RefreshAsync(source); Raise(nameof(WikiIndexStatus));
        var status = ServiceLocator.WikiVocabulary.GetStatus(source); WikiSettingsStatus = status.Status == "success" ? $"Wiki-Suchindex aktualisiert: {status.PageCount:N0} Seiten." : "Wiki-Suchindex konnte nicht aktualisiert werden.";
    }

    private void SaveWebShortcut()
    {
        var editor=SelectedWebShortcut;if(editor==null){WebShortcutStatus="Bitte eine Webseite auswählen.";return;} var model=editor.ToModel();
        if(!Uri.TryCreate(model.Url,UriKind.Absolute,out var uri)||(uri.Scheme!=Uri.UriSchemeHttp&&uri.Scheme!=Uri.UriSchemeHttps)){WebShortcutStatus="Bitte eine gültige absolute http/https URL eingeben.";return;}
        if(model.AutoLogin&&(uri.Scheme!=Uri.UriSchemeHttps||string.IsNullOrWhiteSpace(model.Username))){WebShortcutStatus="Auto Login erfordert HTTPS und einen Benutzernamen.";return;}
        if(editor.Password!=ShortcutPasswordMask)_settings.SetWebShortcutPassword(model,editor.Password); if(model.AutoLogin&&string.IsNullOrWhiteSpace(model.PasswordEncrypted)){WebShortcutStatus="Auto Login erfordert ein Passwort.";return;}
        var i=_settings.Current.WebShortcuts.FindIndex(x=>x.Id==model.Id);if(i<0)_settings.Current.WebShortcuts.Add(model);else _settings.Current.WebShortcuts[i]=model;
        if(!_settings.TrySave()){WebShortcutStatus="Webseite konnte nicht gespeichert werden.";return;} editor.SetEncrypted(model.PasswordEncrypted);editor.Password=model.PasswordEncrypted.Length>0?ShortcutPasswordMask:"";WebShortcutStatus=$"Webseite '{DisplayName(model)}' wurde gespeichert.";
    }
    private void RemoveWebShortcut(){var editor=SelectedWebShortcut;if(editor==null)return;if(MessageBox.Show($"Webseite '{editor.Name}' wirklich entfernen?","Webseite entfernen",MessageBoxButton.YesNo,MessageBoxImage.Warning)!=MessageBoxResult.Yes)return;_settings.Current.WebShortcuts.RemoveAll(x=>x.Id==editor.Id);if(!_settings.TrySave()){WebShortcutStatus="Webseite konnte nicht entfernt werden.";return;}WebShortcuts.Remove(editor);SelectedWebShortcut=WebShortcuts.FirstOrDefault();WebShortcutStatus="Webseite wurde entfernt.";}
    private static string DisplayName(WebShortcutSettings x)=>!string.IsNullOrWhiteSpace(x.Name)?x.Name:Uri.TryCreate(x.Url,UriKind.Absolute,out var u)?u.Host:"Webseite";
    private static string FormatWikiIndexStatus(WikiSourceSettings source)
    {
        var status = ServiceLocator.WikiVocabulary.GetStatus(source); return $"{status.PageCount:N0} Seiten · Zuletzt aktualisiert: {(status.UpdatedUtc?.ToLocalTime().ToString("dd.MM.yyyy HH:mm") ?? "noch nie")}";
    }

    private static string DescribeWikiTestError(Exception exception)
    {
        if (exception is TaskCanceledException) return "Wiki antwortet nicht innerhalb des Zeitlimits.";
        if (exception is JsonException) return "Das Antwortformat des Wikis ist unerwartet.";
        if (exception is HttpRequestException http) return http.StatusCode switch
        {
            HttpStatusCode.Unauthorized => "Authentifizierung fehlgeschlagen (HTTP 401).",
            HttpStatusCode.Forbidden => "Zugriff verweigert (HTTP 403).",
            HttpStatusCode.NotFound => "Wiki API Endpoint nicht gefunden (HTTP 404).",
            _ when http.StatusCode.HasValue => $"Wiki-Anfrage fehlgeschlagen (HTTP {(int)http.StatusCode.Value}).",
            _ => "Wiki-Server ist nicht erreichbar. Bitte DNS, Netzwerk und Base URL prüfen."
        };
        if (exception is UriFormatException) return "Base URL ist ungültig.";
        return exception is InvalidOperationException ? exception.Message : "Wiki-Test fehlgeschlagen. Bitte logs.txt prüfen.";
    }

    private void NotifySettingsConsumers()
    {
        _notifications.HandleSettingsChanged(); _outlookCalendar.HandleSettingsChanged(); _ticketSystem.HandleSettingsChanged(); Raise(string.Empty);
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
