using System.Reflection;
using TaskTool.Models;

namespace TaskTool.Services;

public sealed class AppVersionService
{
    private readonly SettingsService _settings; private readonly LoggerService _logger;
    public string InstalledVersionText => TryGetInstalledVersion(out var version) ? version.ToString() : _settings.Current.InstalledVersion;
    public AppVersionService(SettingsService settings, LoggerService logger) { _settings = settings; _logger = logger; }
    public static SemanticVersion ParseVersion(string value) => SemanticVersion.Parse(value);
    public bool TryGetInstalledVersion(out SemanticVersion version)
    {
        var assemblyText = Assembly.GetEntryAssembly()?.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion;
        if (SemanticVersion.TryParse(assemblyText, out version)) { _logger.Info($"[UpdateVersion] installedVersion={version} source=AssemblyInformationalVersion"); return true; }
        if (SemanticVersion.TryParse(_settings.Current.InstalledVersion, out version)) { _logger.Info($"[UpdateVersion] installedVersion={version} source=LegacySettings"); return true; }
        version = null!; _logger.Error("[UpdateVersion] noValidInstalledVersion=true"); return false;
    }
    public bool ApplyPostUpdateVersion(string targetVersionText)
    {
        if (!SemanticVersion.TryParse(targetVersionText, out var target) || !TryGetInstalledVersion(out var previous) || target < previous) return false;
        var old = _settings.Current.InstalledVersion; _settings.Current.InstalledVersion = target.ToString(); if (_settings.TrySave()) return true; _settings.Current.InstalledVersion = old; return false;
    }
}
