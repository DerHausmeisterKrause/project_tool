using System.Reflection;

namespace TaskTool.Services;

public sealed class AppVersionService
{
    public string CurrentVersionText { get; }
    public Version CurrentVersion { get; }

    public AppVersionService()
    {
        var assembly = typeof(AppVersionService).Assembly;
        var informational = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        CurrentVersionText = (informational?.Split('+')[0] ?? assembly.GetName().Version?.ToString(3) ?? "0.0.0").Trim();
        CurrentVersion = ParseVersion(CurrentVersionText);
    }

    public static Version ParseVersion(string value)
    {
        value = value.Trim();
        if (value.StartsWith('v') || value.StartsWith('V')) value = value[1..];
        var suffix = value.IndexOfAny(['-', '+']);
        if (suffix >= 0) value = value[..suffix];
        return Version.TryParse(value, out var version)
            ? version
            : throw new FormatException($"Ungültige Versionsnummer: {value}");
    }
}
