using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using TaskTool.Models;

namespace TaskTool.Services;

public class SettingsService
{
    [Flags]
    private enum NormalizationChanges
    {
        None = 0,
        TicketUpdateRoute = 1,
        InstalledVersion = 2,
        SegmentDuration = 4,
        LogLevel = 8
    }

    private readonly LoggerService _logger;
    private readonly string _path = Path.Combine(AppContext.BaseDirectory, "settings.json");
    public AppSettings Current { get; private set; } = new();

    public SettingsService(LoggerService logger)
    {
        _logger = logger;
        Load();
        ApplyLoggerMinimumLevel();
    }

    public void Load()
    {
        try
        {
            if (!File.Exists(_path))
            {
                var initialChanges = Normalize(Current);
                LogMigrations(initialChanges);
                Save();
                return;
            }
            var json = File.ReadAllText(_path);
            using var settingsDocument = JsonDocument.Parse(json);
            var hasSegmentDuration = settingsDocument.RootElement.EnumerateObject()
                .Any(property => string.Equals(property.Name, nameof(AppSettings.DefaultSegmentDurationMinutes), StringComparison.OrdinalIgnoreCase));
            Current = JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
            if (!hasSegmentDuration)
                Current.DefaultSegmentDurationMinutes = 0;
            var changes = Normalize(Current);
            if (changes.HasFlag(NormalizationChanges.LogLevel))
                _logger.Warning("Invalid LogLevel in settings.json, fallback='Warning'.");
            if (changes != NormalizationChanges.None)
            {
                LogMigrations(changes);
                Save();
            }
        }
        catch (Exception ex)
        {
            _logger.Error($"Settings load failed: {ex.Message}");
            Current = new AppSettings();
            Normalize(Current);
        }
    }

    private void LogMigrations(NormalizationChanges changes)
    {
        if (changes.HasFlag(NormalizationChanges.TicketUpdateRoute))
            _logger.Info($"[ZnunySettingsMigration] TicketSystemTicketUpdateRoute old='{AppSettings.LegacyTicketSystemTicketUpdateRoute}' new='{AppSettings.DefaultTicketSystemTicketUpdateRoute}'");
        if (changes.HasFlag(NormalizationChanges.InstalledVersion))
            _logger.Info($"[SettingsMigration] InstalledVersion missing initializedVersion={AppSettings.InitialInstalledVersion}");
    }

    private static NormalizationChanges Normalize(AppSettings settings)
    {
        var changes = NormalizationChanges.None;
        if (!Enum.TryParse<AppLogLevel>(settings.LogLevel, true, out var logLevel) || !Enum.IsDefined(logLevel))
        {
            settings.LogLevel = nameof(AppLogLevel.Warning);
            changes |= NormalizationChanges.LogLevel;
        }
        else
        {
            settings.LogLevel = logLevel.ToString();
        }
        var ticketUpdateRouteMigrated = string.Equals(
            settings.TicketSystemTicketUpdateRoute?.Trim(),
            AppSettings.LegacyTicketSystemTicketUpdateRoute,
            StringComparison.OrdinalIgnoreCase);
        if (ticketUpdateRouteMigrated)
            changes |= NormalizationChanges.TicketUpdateRoute;
        if (string.IsNullOrWhiteSpace(settings.InstalledVersion))
        {
            settings.InstalledVersion = AppSettings.InitialInstalledVersion;
            changes |= NormalizationChanges.InstalledVersion;
        }
        settings.CurrentTasksSortField = string.Equals(settings.CurrentTasksSortField, "Created", StringComparison.OrdinalIgnoreCase)
            ? "Created"
            : "Updated";
        if (settings.DefaultSegmentDurationMinutes is < 15 or > 240)
        {
            settings.DefaultSegmentDurationMinutes = 30;
            changes |= NormalizationChanges.SegmentDuration;
        }
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

        if (!string.Equals(settings.OutlookCalendarSyncMode, "Manual", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(settings.OutlookCalendarSyncMode, "Periodic", StringComparison.OrdinalIgnoreCase))
            settings.OutlookCalendarSyncMode = "Manual";

        settings.OutlookCalendarSyncIntervalMinutes = Math.Clamp(settings.OutlookCalendarSyncIntervalMinutes, 1, 60);
        settings.OutlookCalendarRangePastDays = settings.OutlookCalendarRangePastDays <= 0 ? 14 : Math.Clamp(settings.OutlookCalendarRangePastDays, 1, 30);
        settings.OutlookCalendarRangeFutureDays = settings.OutlookCalendarRangeFutureDays <= 0 ? 14 : Math.Clamp(settings.OutlookCalendarRangeFutureDays, 1, 90);
        var supportedCalendarTimeZones = new[] { "Europe/Berlin", "Europe/London", "UTC", "Europe/Vienna", "Europe/Zurich" };
        if (!supportedCalendarTimeZones.Contains(settings.CalendarTimeZoneId, StringComparer.Ordinal))
            settings.CalendarTimeZoneId = "Europe/Berlin";

        settings.TicketSystemWebUrl = settings.TicketSystemWebUrl?.Trim() ?? "https://SERVER/index.pl";
        if (string.IsNullOrWhiteSpace(settings.TicketSystemWebUrl))
            settings.TicketSystemWebUrl = "https://SERVER/index.pl";
        settings.TicketSystemApiUrl = settings.TicketSystemApiUrl?.Trim() ?? "https://SERVER/nph-genericinterface.pl/Webservice/GenericTicketConnectorREST";
        if (string.IsNullOrWhiteSpace(settings.TicketSystemApiUrl))
            settings.TicketSystemApiUrl = "https://SERVER/nph-genericinterface.pl/Webservice/GenericTicketConnectorREST";
        settings.TicketSystemUsername = settings.TicketSystemUsername?.Trim() ?? string.Empty;
        settings.TicketSystemPasswordEncrypted ??= string.Empty;
        settings.TicketSystemPassword ??= string.Empty;
        settings.TicketSystemAgentId = Math.Max(0, settings.TicketSystemAgentId);
        settings.TicketSystemCandidateUserId = Math.Max(1, settings.TicketSystemCandidateUserId);
        settings.TicketSystemCandidateKeywords = settings.TicketSystemCandidateKeywords?.Trim() ?? string.Empty;
        settings.TicketSystemTicketSearchRoute = NormalizeRoute(settings.TicketSystemTicketSearchRoute, "/Ticket");
        settings.TicketSystemTicketSearchMethod = string.Equals(settings.TicketSystemTicketSearchMethod, "POST", StringComparison.OrdinalIgnoreCase) ? "POST" : "GET";
        settings.TicketSystemTicketSearchAuthMode = string.Equals(settings.TicketSystemTicketSearchAuthMode, "Direct", StringComparison.OrdinalIgnoreCase) ? "Direct" : "Session";
        settings.TicketSystemTicketGetRouteTemplate = NormalizeRoute(settings.TicketSystemTicketGetRouteTemplate, "/Ticket/{TicketID}");
        settings.TicketSystemTicketGetMethod = "GET";
        settings.TicketSystemTicketGetAuthMode = string.Equals(settings.TicketSystemTicketGetAuthMode, "Direct", StringComparison.OrdinalIgnoreCase) ? "Direct" : "Session";
        settings.TicketSystemTicketUpdateRoute = ticketUpdateRouteMigrated
            ? AppSettings.DefaultTicketSystemTicketUpdateRoute
            : NormalizeRoute(settings.TicketSystemTicketUpdateRoute, AppSettings.DefaultTicketSystemTicketUpdateRoute);
        settings.TicketSystemTicketCreateRoute = NormalizeRoute(settings.TicketSystemTicketCreateRoute, "/Ticket");
        settings.TicketSystemTicketCreateMethod = string.IsNullOrWhiteSpace(settings.TicketSystemTicketCreateMethod)
            ? "POST"
            : settings.TicketSystemTicketCreateMethod.Trim().ToUpperInvariant();
        settings.TicketSystemCreateQueue = settings.TicketSystemCreateQueue?.Trim() ?? string.Empty;
        settings.TicketSystemCreateState = string.IsNullOrWhiteSpace(settings.TicketSystemCreateState) ? "open" : settings.TicketSystemCreateState.Trim();
        settings.TicketSystemCreatePriority = string.IsNullOrWhiteSpace(settings.TicketSystemCreatePriority) ? "3 normal" : settings.TicketSystemCreatePriority.Trim();
        settings.TicketSystemCreateType = settings.TicketSystemCreateType?.Trim() ?? string.Empty;
        settings.TicketSystemCreateCustomerUser = settings.TicketSystemCreateCustomerUser?.Trim() ?? string.Empty;
        settings.TicketSystemDynamicFieldOptionsRoute = string.Equals(settings.TicketSystemDynamicFieldOptionsRoute?.Trim(), "/DynamicField/Options", StringComparison.OrdinalIgnoreCase)
            ? "/Ticket/DynamicField/{FieldName}/Options"
            : NormalizeRoute(settings.TicketSystemDynamicFieldOptionsRoute, "/Ticket/DynamicField/{FieldName}/Options");
        settings.TicketSystemCostCenterFieldName = string.IsNullOrWhiteSpace(settings.TicketSystemCostCenterFieldName) ? "KostenstelleID" : settings.TicketSystemCostCenterFieldName.Trim();
        settings.TicketSystemOrderFieldName = string.IsNullOrWhiteSpace(settings.TicketSystemOrderFieldName) ? "AuftragsID" : settings.TicketSystemOrderFieldName.Trim();
        settings.TicketSystemCostCenterOptions = settings.TicketSystemCostCenterOptions?.Trim() ?? string.Empty;
        settings.TicketSystemOrderOptions = settings.TicketSystemOrderOptions?.Trim() ?? string.Empty;
        settings.HomeOfficeMailRecipient1 = settings.HomeOfficeMailRecipient1?.Trim() ?? string.Empty;
        settings.HomeOfficeMailRecipient2 = settings.HomeOfficeMailRecipient2?.Trim() ?? string.Empty;
        settings.TicketSystemSyncIntervalMinutes = settings.TicketSystemSyncIntervalMinutes <= 0 ? 15 : Math.Clamp(settings.TicketSystemSyncIntervalMinutes, 1, 1440);
        if (!settings.TicketSystemIncludeOwner && !settings.TicketSystemIncludeResponsible)
            settings.TicketSystemIncludeOwner = true;
        if (settings.TicketSystemOnlyOpenTickets)
            settings.TicketSystemShowClosedTickets = false;
        if (!settings.TicketSystemAutofillCredentials)
            settings.TicketSystemAutoLogin = false;
        return changes;
    }

    private static string NormalizeRoute(string? route, string defaultRoute)
    {
        if (string.IsNullOrWhiteSpace(route)) return defaultRoute;
        route = route.Trim();
        return route.StartsWith('/') ? route : "/" + route;
    }

    public string GetTicketSystemPassword()
    {
        if (!string.IsNullOrWhiteSpace(Current.TicketSystemPasswordEncrypted))
            return Unprotect(Current.TicketSystemPasswordEncrypted);

        if (string.IsNullOrEmpty(Current.TicketSystemPassword))
            return string.Empty;

        var migratedPassword = Current.TicketSystemPassword;
        Current.TicketSystemPasswordEncrypted = Protect(migratedPassword);
        Current.TicketSystemPassword = string.Empty;
        Save();
        return migratedPassword;
    }

    public void SetTicketSystemPassword(string password)
    {
        Current.TicketSystemPasswordEncrypted = string.IsNullOrEmpty(password) ? string.Empty : Protect(password);
        Current.TicketSystemPassword = string.Empty;
    }

    public void Save()
    {
        TrySave();
    }

    public bool TrySave()
    {
        try
        {
            ApplyLoggerMinimumLevel();
            var json = JsonSerializer.Serialize(Current, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(_path, json);
            SettingsChanged?.Invoke();
            return true;
        }
        catch (Exception ex)
        {
            _logger.Error($"Settings save failed: {ex.Message}");
            return false;
        }
    }

    public event Action? SettingsChanged;

    private void ApplyLoggerMinimumLevel()
    {
        var level = Enum.TryParse<AppLogLevel>(Current.LogLevel, true, out var parsed) && Enum.IsDefined(parsed)
            ? parsed
            : AppLogLevel.Warning;
        Current.LogLevel = level.ToString();
        _logger.SetMinimumLevel(level);
    }

    private string Protect(string value)
    {
        try
        {
            var bytes = Encoding.UTF8.GetBytes(value);
            var protectedBytes = ProtectedData.Protect(bytes, null, DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(protectedBytes);
        }
        catch (Exception ex)
        {
            _logger.Error($"Settings password encryption failed: {ex.Message}");
            return string.Empty;
        }
    }

    private string Unprotect(string encryptedValue)
    {
        try
        {
            var protectedBytes = Convert.FromBase64String(encryptedValue);
            var bytes = ProtectedData.Unprotect(protectedBytes, null, DataProtectionScope.CurrentUser);
            return Encoding.UTF8.GetString(bytes);
        }
        catch (Exception ex)
        {
            _logger.Error($"Settings password decryption failed: {ex.Message}");
            return string.Empty;
        }
    }
}
