namespace TaskTool.Services;

public sealed class AppVersionService
{
    private readonly SettingsService _settings;
    private readonly LoggerService _logger;
    public string InstalledVersionText => _settings.Current.InstalledVersion;

    public AppVersionService(SettingsService settings, LoggerService logger)
    {
        _settings = settings;
        _logger = logger;
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

    public bool TryGetInstalledVersion(out Version version)
    {
        try
        {
            version = ParseVersion(InstalledVersionText);
            _logger.Info($"[UpdateVersion] installedVersion={version} source=Settings");
            return true;
        }
        catch (Exception)
        {
            version = new Version();
            _logger.Error($"[UpdateVersion] invalidLocalVersion='{InstalledVersionText}'");
            return false;
        }
    }

    public bool ApplyPostUpdateVersion(string targetVersionText)
    {
        Version targetVersion;
        try { targetVersion = ParseVersion(targetVersionText); }
        catch (Exception)
        {
            _logger.Error($"[PostUpdate] invalidTargetVersion='{targetVersionText}'");
            return false;
        }

        if (!TryGetInstalledVersion(out var previousVersion)) return false;
        if (targetVersion < previousVersion)
        {
            _logger.Error($"[PostUpdate] previousVersion={previousVersion} newVersion={targetVersion} settingsUpdated=false reason=downgrade");
            return false;
        }

        var previousVersionText = _settings.Current.InstalledVersion;
        _settings.Current.InstalledVersion = targetVersion.ToString();
        if (!_settings.TrySave())
        {
            _settings.Current.InstalledVersion = previousVersionText;
            _logger.Error($"[PostUpdate] previousVersion={previousVersion} newVersion={targetVersion} settingsUpdated=false reason=settings-save-failed");
            return false;
        }
        _logger.Info($"[PostUpdate] previousVersion={previousVersion} newVersion={targetVersion} settingsUpdated=true");
        return true;
    }
}
