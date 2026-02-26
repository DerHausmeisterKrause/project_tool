using TaskTool.ViewModels;

namespace TaskTool.Services;

public static class ServiceLocator
{
    public static LoggerService Logger { get; private set; } = null!;
    public static SettingsService Settings { get; private set; } = null!;
    public static DatabaseService Database { get; private set; } = null!;
    public static OutlookInteropService Outlook { get; private set; } = null!;
    public static NotificationService Notifications { get; private set; } = null!;
    public static TaskService Tasks { get; private set; } = null!;
    public static WorkDayService WorkDays { get; private set; } = null!;
    public static MainViewModel MainViewModel { get; private set; } = null!;

    public static void Initialize()
    {
        Logger = new LoggerService();
        Settings = new SettingsService(Logger);
        Database = new DatabaseService(Logger);
        Database.Initialize();
        Outlook = new OutlookInteropService(Logger, Settings);
        Tasks = new TaskService(Database, Logger, Outlook, Settings);
        WorkDays = new WorkDayService(Database, Logger);
        Notifications = new NotificationService(Logger, Settings, Tasks);
        MainViewModel = new MainViewModel(Tasks, WorkDays, Settings, Notifications, Logger);
    }
}
