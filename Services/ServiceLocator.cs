using TaskTool.ViewModels;

namespace TaskTool.Services;

public static class ServiceLocator
{
    public static LoggerService Logger { get; private set; } = null!;
    public static SettingsService Settings { get; private set; } = null!;
    public static DatabaseService Database { get; private set; } = null!;
    public static OutlookInteropService Outlook { get; private set; } = null!;
    public static NotificationService Notifications { get; private set; } = null!;
    public static OutlookCalendarService OutlookCalendar { get; private set; } = null!;
    public static TaskService Tasks { get; private set; } = null!;
    public static WorkDayService WorkDays { get; private set; } = null!;
    public static TicketSystemService TicketSystem { get; private set; } = null!;
    public static GermanTimeService GermanTime { get; private set; } = null!;
    public static AppVersionService AppVersion { get; private set; } = null!;
    public static UpdateService Updates { get; private set; } = null!;
    public static HomeOfficeService HomeOffice { get; private set; } = null!;
    public static TicketAssignmentSnapshotService TicketAssignmentSnapshots { get; private set; } = null!;
    public static TicketCandidateSnapshotService TicketCandidateSnapshots { get; private set; } = null!;
    public static TicketCandidateScanStateService TicketCandidateScanStates { get; private set; } = null!;
    public static TicketDetailCacheService TicketDetails { get; private set; } = null!;
    public static WikiSearchService WikiSearch { get; private set; } = null!;
    public static WikiVocabularyIndexService WikiVocabulary { get; private set; } = null!;
    public static WebShortcutBrowserSessionManager WebShortcutBrowsers { get; private set; } = null!;
    public static MainViewModel MainViewModel { get; private set; } = null!;

    public static void Initialize()
    {
        Logger = new LoggerService(AppLogLevel.Warning);
        Settings = new SettingsService(Logger);
        Logger.SetMinimumLevel(ParseLogLevel(Settings.Current.LogLevel));
        AppVersion = new AppVersionService(Settings, Logger);
        Updates = new UpdateService(Logger, Settings, AppVersion);
        GermanTime = new GermanTimeService();
        Database = new DatabaseService(Logger);
        Database.Initialize();
        WikiVocabulary = new WikiVocabularyIndexService(Database, Settings, Logger);
        WikiSearch = new WikiSearchService(Database, Settings, Logger, vocabulary: WikiVocabulary);
        WebShortcutBrowsers = new WebShortcutBrowserSessionManager(Settings, Logger);
        _ = WikiVocabulary.RefreshStaleAsync();
        Outlook = new OutlookInteropService(Logger, Settings);
        Tasks = new TaskService(Database, Logger, Outlook, Settings);
        TicketAssignmentSnapshots = new TicketAssignmentSnapshotService(Database);
        TicketCandidateSnapshots = new TicketCandidateSnapshotService(Database);
        TicketCandidateScanStates = new TicketCandidateScanStateService(Database);
        TicketDetails = new TicketDetailCacheService(Database);
        WorkDays = new WorkDayService(Database, Logger);
        Notifications = new NotificationService(Logger, Settings, Tasks);
        TicketSystem = new TicketSystemService(Settings, Tasks, TicketAssignmentSnapshots, TicketCandidateSnapshots, TicketCandidateScanStates, TicketDetails, Notifications, Logger);
        OutlookCalendar = new OutlookCalendarService(Logger, Settings, Outlook, WorkDays);
        HomeOffice = new HomeOfficeService(WorkDays, Settings, Outlook, OutlookCalendar, Logger);
        MainViewModel = new MainViewModel(Tasks, WorkDays, Settings, Notifications, OutlookCalendar, TicketSystem, Updates, HomeOffice, GermanTime, Logger);
    }

    private static AppLogLevel ParseLogLevel(string? value)
        => Enum.TryParse<AppLogLevel>(value, true, out var level) && Enum.IsDefined(level)
            ? level
            : AppLogLevel.Warning;
}
