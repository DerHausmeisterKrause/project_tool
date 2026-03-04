using System.IO;
using System.Text.Json;
using TaskTool.Models;

namespace TaskTool.Services;

public class SettingsService
{
    private readonly LoggerService _logger;
    private readonly string _path = Path.Combine(AppContext.BaseDirectory, "settings.json");
    public AppSettings Current { get; private set; } = new();

    public SettingsService(LoggerService logger)
    {
        _logger = logger;
        Load();
    }

    public void Load()
    {
        try
        {
            if (!File.Exists(_path))
            {
                Normalize(Current);
                Save();
                return;
            }
            var json = File.ReadAllText(_path);
            Current = JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
            Normalize(Current);
        }
        catch (Exception ex)
        {
            _logger.Error($"Settings load failed: {ex.Message}");
            Current = new AppSettings();
            Normalize(Current);
        }
    }

    private static void Normalize(AppSettings settings)
    {
        if (settings.FridayTargetMinutes <= 0)
            settings.FridayTargetMinutes = 300;

        if (string.IsNullOrWhiteSpace(settings.DynamicIslandDockPosition))
            settings.DynamicIslandDockPosition = "TopCenter";

        var validDockPositions = new[]
        {
            "TopCenter", "TopLeft", "TopRight", "LeftCenter", "RightCenter", "BottomLeft", "BottomCenter", "BottomRight"
        };

        if (!validDockPositions.Contains(settings.DynamicIslandDockPosition, StringComparer.OrdinalIgnoreCase))
            settings.DynamicIslandDockPosition = "TopCenter";
    }

    public void Save()
    {
        try
        {
            var json = JsonSerializer.Serialize(Current, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_path, json);
        }
        catch (Exception ex)
        {
            _logger.Error($"Settings save failed: {ex.Message}");
        }
    }
}
