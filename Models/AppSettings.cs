namespace TaskTool.Models;

public class AppSettings
{
    // Bootstrap/delivery version for fresh settings created by this distribution.
    // Existing non-empty InstalledVersion values remain authoritative until a successful update.
    public const string InitialInstalledVersion = "2.1.0";
    public const string DefaultTicketSystemTicketUpdateRoute = "/Ticket/{TicketID}/Update";
    public const string LegacyTicketSystemTicketUpdateRoute = "/Ticket/Update";
    public const string DefaultTicketSystemTicketCreateRoute = "/TicketCreate";
    public const string LegacyTicketSystemTicketCreateRoute = "/Ticket";
    public bool DynamicIslandEnabled { get; set; } = true;
    public string LogLevel { get; set; } = "Warning";
    public bool NotificationSoundEnabled { get; set; } = true;
    public bool CheckForUpdatesOnStartup { get; set; } = true;
    // Retained for old settings.json files only; silent installation is never authorized by this value.
    public bool AutoInstallUpdatesOnStartup { get; set; } = false;
    public string UpdateChannel { get; set; } = "Stable";
    public string InstalledVersion { get; set; } = string.Empty;
    public string CurrentTasksSortField { get; set; } = "Updated";
    public bool CurrentTasksSortDescending { get; set; } = true;
    public bool HidePastTodayItems { get; set; } = true;
    public int DefaultSegmentDurationMinutes { get; set; } = 30;
    public string HomeOfficeMailRecipient1 { get; set; } = string.Empty;
    public string HomeOfficeMailRecipient2 { get; set; } = string.Empty;
    public string DynamicIslandDockPosition { get; set; } = "TopCenter";

    public bool OutlookSyncEnabled { get; set; } = true;
    public string OutlookCategoryName { get; set; } = "FocusBlock";

    public bool OutlookCalendarEnabled { get; set; } = false;
    public bool OutlookConflictWarningsEnabled { get; set; } = true;
    public bool OutlookTeamsButtonEnabled { get; set; } = true;
    public string OutlookCalendarSyncMode { get; set; } = "Manual";
    public int OutlookCalendarSyncIntervalMinutes { get; set; } = 5;
    public int OutlookCalendarRangePastDays { get; set; } = 14;
    public int OutlookCalendarRangeFutureDays { get; set; } = 14;
    public bool OutlookInterpretAllDayAsMarkers { get; set; } = true;
    public bool ShowWeekendInWeekView { get; set; } = false;
    public bool ShowInternalTaskSegmentsInCalendar { get; set; } = false;
    public string CalendarTimeZoneId { get; set; } = "Europe/Berlin";
    public int ReminderLeadMinutes { get; set; } = 2;
    public string DateTimeFormat { get; set; } = "yyyy-MM-dd HH:mm";

    public string TicketSystemWebUrl { get; set; } = "https://SERVER/index.pl";
    public string TicketSystemApiUrl { get; set; } = "https://SERVER/nph-genericinterface.pl/Webservice/GenericTicketConnectorREST";
    public string TicketSystemUsername { get; set; } = string.Empty;
    public string TicketSystemPasswordEncrypted { get; set; } = string.Empty;
    public string TicketSystemPassword { get; set; } = string.Empty;
    public int TicketSystemAgentId { get; set; } = 0;
    public int TicketSystemCandidateUserId { get; set; } = 1;
    public bool TicketSystemHidePendingTickets { get; set; } = false;
    public bool TicketSystemNotifyPendingTickets { get; set; } = true;
    public string TicketSystemCandidateKeywords { get; set; } = string.Empty;
    public string TicketSystemCandidateExcludeKeywords { get; set; } = string.Empty;
    public string TicketSystemTicketSearchRoute { get; set; } = "/Ticket";
    public string TicketSystemTicketSearchMethod { get; set; } = "GET";
    public string TicketSystemTicketSearchAuthMode { get; set; } = "Session";
    public string TicketSystemTicketGetRouteTemplate { get; set; } = "/Ticket/{TicketID}";
    public string TicketSystemTicketGetMethod { get; set; } = "GET";
    public string TicketSystemTicketGetAuthMode { get; set; } = "Session";
    public string TicketSystemTicketUpdateRoute { get; set; } = DefaultTicketSystemTicketUpdateRoute;
    public string TicketSystemTicketCreateRoute { get; set; } = DefaultTicketSystemTicketCreateRoute;
    public string TicketSystemTicketCreateMethod { get; set; } = "POST";
    public string TicketSystemCreateQueue { get; set; } = string.Empty;
    public string TicketSystemCreateState { get; set; } = "open";
    public string TicketSystemCreatePriority { get; set; } = "3 normal";
    public string TicketSystemCreateType { get; set; } = string.Empty;
    public string TicketSystemCreateCustomerUser { get; set; } = string.Empty;
    public string TicketSystemReplyTemplate { get; set; } = string.Empty;
    public string TicketSystemDynamicFieldOptionsRoute { get; set; } = "/Ticket/DynamicField/{FieldName}/Options";
    public string TicketSystemCostCenterFieldName { get; set; } = "KostenstelleID";
    public string TicketSystemOrderFieldName { get; set; } = "AuftragsID";
    public string TicketSystemCostCenterOptions { get; set; } = string.Empty;
    public string TicketSystemOrderOptions { get; set; } = string.Empty;
    public int TicketSystemSyncIntervalMinutes { get; set; } = 15;
    public bool TicketSystemOnlyOpenTickets { get; set; } = true;
    public bool TicketSystemShowClosedTickets { get; set; } = false;
    public bool TicketSystemIncludeOwner { get; set; } = true;
    public bool TicketSystemIncludeResponsible { get; set; } = true;
    public bool NotifyOnNewAssignedTickets { get; set; } = true;
    public bool TicketSystemAutofillCredentials { get; set; } = false;
    public bool TicketSystemAutoLogin { get; set; } = false;
    public List<WikiSourceSettings> WikiSources { get; set; } = new();
    public string DefaultWikiSourceId { get; set; } = string.Empty;
    public List<WebShortcutSettings> WebShortcuts { get; set; } = new();

    public int MondayTargetMinutes { get; set; } = 480;
    public int TuesdayTargetMinutes { get; set; } = 480;
    public int WednesdayTargetMinutes { get; set; } = 480;
    public int ThursdayTargetMinutes { get; set; } = 480;
    public int FridayTargetMinutes { get; set; } = 300;
    public int SaturdayTargetMinutes { get; set; } = 0;
    public int SundayTargetMinutes { get; set; } = 0;

    public int GetTargetMinutes(DayOfWeek day) => day switch
    {
        DayOfWeek.Monday => MondayTargetMinutes,
        DayOfWeek.Tuesday => TuesdayTargetMinutes,
        DayOfWeek.Wednesday => WednesdayTargetMinutes,
        DayOfWeek.Thursday => ThursdayTargetMinutes,
        DayOfWeek.Friday => FridayTargetMinutes,
        DayOfWeek.Saturday => SaturdayTargetMinutes,
        DayOfWeek.Sunday => SundayTargetMinutes,
        _ => 0
    };
}

public sealed class WebShortcutSettings
{
    public string Id { get; set; } = Guid.NewGuid().ToString(); public string Name { get; set; } = string.Empty; public string Url { get; set; } = string.Empty;
    public string Username { get; set; } = string.Empty; public string PasswordEncrypted { get; set; } = string.Empty; public bool Enabled { get; set; } = true; public bool AutoLogin { get; set; }
    public bool DisableWebSecurity { get; set; } = false;
    public bool AllowInsecureContent { get; set; } = false;
    public int SortOrder { get; set; }
}

public sealed class WikiSourceSettings
{
    public string Id { get; set; } = Guid.NewGuid().ToString();
    public string Name { get; set; } = string.Empty;
    public bool Enabled { get; set; } = true;
    public string ProviderType { get; set; } = "ConfluenceDataCenter";
    public string BaseUrl { get; set; } = string.Empty;
    public string AuthMode { get; set; } = "BearerToken";
    public string Username { get; set; } = string.Empty;
    public string SecretEncrypted { get; set; } = string.Empty;
    public string ApiKeyHeaderName { get; set; } = "X-API-Key";
    public string SpaceKey { get; set; } = string.Empty;
    public bool SearchAllSpaces { get; set; } = true;
    public List<string> SpaceKeys { get; set; } = new();
    public string BrowserHomeUrl { get; set; } = string.Empty;
    public string BrowserLoginMode { get; set; } = "BrowserSession";
    public string BrowserUsername { get; set; } = string.Empty;
    public string BrowserPasswordEncrypted { get; set; } = string.Empty;
    public bool BrowserAutoSubmit { get; set; } = true;
    public int MaxResults { get; set; } = 5;
    public string HttpMethod { get; set; } = "GET";
    public string SearchUrlTemplate { get; set; } = string.Empty;
    public string RequestBodyTemplate { get; set; } = string.Empty;
    public string ResultArrayPath { get; set; } = "$.results";
    public string ResultIdPath { get; set; } = "id";
    public string ResultTitlePath { get; set; } = "title";
    public string ResultUrlPath { get; set; } = "url";
    public string ResultExcerptPath { get; set; } = "excerpt";
}
