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

        if (!string.Equals(settings.OutlookCalendarSyncMode, "Manual", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(settings.OutlookCalendarSyncMode, "Periodic", StringComparison.OrdinalIgnoreCase))
            settings.OutlookCalendarSyncMode = "Manual";

        settings.OutlookCalendarSyncIntervalMinutes = Math.Clamp(settings.OutlookCalendarSyncIntervalMinutes, 1, 60);
        settings.OutlookCalendarRangePastDays = settings.OutlookCalendarRangePastDays <= 0 ? 14 : Math.Clamp(settings.OutlookCalendarRangePastDays, 1, 30);
        settings.OutlookCalendarRangeFutureDays = settings.OutlookCalendarRangeFutureDays <= 0 ? 14 : Math.Clamp(settings.OutlookCalendarRangeFutureDays, 1, 90);

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
        settings.TicketSystemTicketSearchRoute = NormalizeRoute(settings.TicketSystemTicketSearchRoute, "/Ticket");
        settings.TicketSystemTicketSearchMethod = string.Equals(settings.TicketSystemTicketSearchMethod, "POST", StringComparison.OrdinalIgnoreCase) ? "POST" : "GET";
        settings.TicketSystemTicketSearchAuthMode = string.Equals(settings.TicketSystemTicketSearchAuthMode, "Direct", StringComparison.OrdinalIgnoreCase) ? "Direct" : "Session";
        settings.TicketSystemTicketGetRouteTemplate = NormalizeRoute(settings.TicketSystemTicketGetRouteTemplate, "/Ticket/{TicketID}");
        settings.TicketSystemTicketGetMethod = "GET";
        settings.TicketSystemTicketGetAuthMode = string.Equals(settings.TicketSystemTicketGetAuthMode, "Direct", StringComparison.OrdinalIgnoreCase) ? "Direct" : "Session";
        settings.TicketSystemSyncIntervalMinutes = settings.TicketSystemSyncIntervalMinutes <= 0 ? 15 : Math.Clamp(settings.TicketSystemSyncIntervalMinutes, 1, 1440);
        if (!settings.TicketSystemIncludeOwner && !settings.TicketSystemIncludeResponsible)
            settings.TicketSystemIncludeOwner = true;
        if (settings.TicketSystemOnlyOpenTickets)
            settings.TicketSystemShowClosedTickets = false;
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
