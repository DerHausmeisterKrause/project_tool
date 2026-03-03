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

    private void LoadThemeSafely()
    {
        try
        {
            Resources.MergedDictionaries.Add(new ResourceDictionary
            {
                Source = new Uri("Themes/Theme.xaml", UriKind.Relative)
            });
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
