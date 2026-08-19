using System.IO;
using System.Text;

namespace TaskTool.Services;

public enum AppLogLevel
{
    Info = 0,
    Warning = 1,
    Error = 2
}

public class LoggerService
{
    private readonly string _logPath = Path.Combine(AppContext.BaseDirectory, "logs.txt");
    private readonly object _sync = new();
    private int _minimumLevel;

    public LoggerService(AppLogLevel minimumLevel = AppLogLevel.Warning)
    {
        _minimumLevel = (int)minimumLevel;
    }

    public AppLogLevel MinimumLevel => (AppLogLevel)Volatile.Read(ref _minimumLevel);

    public void SetMinimumLevel(AppLogLevel minimumLevel)
        => Interlocked.Exchange(ref _minimumLevel, (int)minimumLevel);

    public void Info(string message) => Write(AppLogLevel.Info, message);
    public void Warning(string message) => Write(AppLogLevel.Warning, message);
    public void Error(string message) => Write(AppLogLevel.Error, message);

    private void Write(AppLogLevel level, string message)
    {
        if (level < MinimumLevel)
            return;

        lock (_sync)
        {
            var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level.ToString().ToUpperInvariant()}] {message}{Environment.NewLine}";
            File.AppendAllText(_logPath, line, Encoding.UTF8);
        }
    }
}
