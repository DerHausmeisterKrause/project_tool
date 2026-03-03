using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using TaskTool.Services;

namespace TaskTool;

public partial class App : Application
{
    private LoggerService? _startupLogger;

    protected override void OnStartup(StartupEventArgs e)
    {
        _startupLogger = new LoggerService();

        DispatcherUnhandledException += OnDispatcherUnhandledException;
        AppDomain.CurrentDomain.UnhandledException += OnCurrentDomainUnhandledException;
        TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;

        try
        {
            LoadThemeSafely();
            ServiceLocator.Initialize();

            var mainWindow = new MainWindow();
            MainWindow = mainWindow;
            ServiceLocator.Notifications.AttachMainWindow(mainWindow);
            mainWindow.Show();
        }
        catch (Exception ex)
        {
            _startupLogger.Error($"Fatal startup error: {ex}");
            MessageBox.Show(
                $"Die Anwendung konnte nicht gestartet werden. Details stehen in logs.txt.\n\n{ex.Message}",
                "TaskTool Startfehler",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
            Shutdown(-1);
            return;
        }

        base.OnStartup(e);
    }


    protected override void OnExit(ExitEventArgs e)
    {
        try
        {
            if (ServiceLocator.Notifications != null)
                ServiceLocator.Notifications.Dispose();
        }
        catch
        {
            // ignore shutdown exceptions
        }

        base.OnExit(e);
    }

    private void LoadThemeSafely()
    {
        try
        {
            var themeUri = new Uri("Themes/Theme.xaml", UriKind.Relative);
            for (var i = Resources.MergedDictionaries.Count - 1; i >= 0; i--)
            {
                var source = Resources.MergedDictionaries[i].Source;
                if (source != null && source.OriginalString.Equals(themeUri.OriginalString, StringComparison.OrdinalIgnoreCase))
                {
                    Resources.MergedDictionaries.RemoveAt(i);
                }
            }

            Resources.MergedDictionaries.Add(new ResourceDictionary { Source = themeUri });
        }
        catch (Exception ex)
        {
            _startupLogger?.Error($"Theme load failed: {ex}");
            MessageBox.Show(
                $"Das Theme konnte nicht geladen werden und wurde übersprungen.\n\n{ex.Message}",
                "TaskTool Theme-Warnung",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }
    }

    private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
    {
        _startupLogger?.Error($"Dispatcher unhandled exception: {e.Exception}");
        MessageBox.Show(
            "Ein unerwarteter Fehler ist aufgetreten. Details stehen in logs.txt.",
            "TaskTool Fehler",
            MessageBoxButton.OK,
            MessageBoxImage.Error);
        e.Handled = true;
    }

    private void OnCurrentDomainUnhandledException(object? sender, UnhandledExceptionEventArgs e)
    {
        if (e.ExceptionObject is Exception ex)
        {
            _startupLogger?.Error($"AppDomain unhandled exception: {ex}");
        }
        else
        {
            _startupLogger?.Error("AppDomain unhandled exception: unknown exception object.");
        }
    }

    private void OnUnobservedTaskException(object? sender, UnobservedTaskExceptionEventArgs e)
    {
        _startupLogger?.Error($"Unobserved task exception: {e.Exception}");
        e.SetObserved();
    }
}
