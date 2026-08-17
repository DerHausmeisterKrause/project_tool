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
    public static MainViewModel MainViewModel { get; private set; } = null!;

    public static void Initialize()
    {
        Logger = new LoggerService();
        Settings = new SettingsService(Logger);
        AppVersion = new AppVersionService(Settings, Logger);
        Updates = new UpdateService(Logger, Settings, AppVersion);
        GermanTime = new GermanTimeService();
        Database = new DatabaseService(Logger);
        Database.Initialize();
        Outlook = new OutlookInteropService(Logger, Settings);
        Tasks = new TaskService(Database, Logger, Outlook, Settings);
        WorkDays = new WorkDayService(Database, Logger);
        TicketSystem = new TicketSystemService(Settings, Tasks, Logger);
        Notifications = new NotificationService(Logger, Settings, Tasks);
        OutlookCalendar = new OutlookCalendarService(Logger, Settings, Outlook);
        HomeOffice = new HomeOfficeService(WorkDays, Settings, Outlook, OutlookCalendar, Logger);
        MainViewModel = new MainViewModel(Tasks, WorkDays, Settings, Notifications, OutlookCalendar, TicketSystem, Updates, HomeOffice, Logger);
    }
}
