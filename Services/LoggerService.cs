using System.IO;
using System.Text;

namespace TaskTool.Services;

public class LoggerService
{
    private readonly string _logPath = Path.Combine(AppContext.BaseDirectory, "logs.txt");
    private readonly object _sync = new();

    public void Info(string message) => Write("INFO", message);
    public void Error(string message) => Write("ERROR", message);

    private void Write(string level, string message)
    {
        lock (_sync)
        {
            var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}] {message}{Environment.NewLine}";
            File.AppendAllText(_logPath, line, Encoding.UTF8);
        }
    }
}
