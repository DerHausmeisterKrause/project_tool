namespace TaskTool.Services;

public sealed class GermanTimeService
{
    private static readonly IReadOnlyDictionary<string, string[]> TimeZoneMappings = new Dictionary<string, string[]>(StringComparer.Ordinal)
    {
        ["Europe/Berlin"] = ["Europe/Berlin", "W. Europe Standard Time"],
        ["Europe/Vienna"] = ["Europe/Vienna", "W. Europe Standard Time"],
        ["Europe/Zurich"] = ["Europe/Zurich", "W. Europe Standard Time"],
        ["Europe/London"] = ["Europe/London", "GMT Standard Time"],
        ["UTC"] = ["UTC"]
    };

    public DateTimeOffset GetLocalNow(string configuredTimeZoneId, DateTimeOffset? utcNow = null)
    {
        var utc = (utcNow ?? DateTimeOffset.UtcNow).ToUniversalTime();
        return TimeZoneInfo.ConvertTime(utc, ResolveTimeZone(configuredTimeZoneId));
    }

    public TimeZoneInfo ResolveTimeZone(string configuredTimeZoneId)
    {
        var stableId = TimeZoneMappings.ContainsKey(configuredTimeZoneId) ? configuredTimeZoneId : "Europe/Berlin";
        foreach (var id in TimeZoneMappings[stableId])
        {
            try
            {
                return TimeZoneInfo.FindSystemTimeZoneById(id);
            }
            catch (TimeZoneNotFoundException)
            {
            }
            catch (InvalidTimeZoneException)
            {
            }
        }

        throw new TimeZoneNotFoundException($"Die Kalender-Zeitzone '{stableId}' ist auf diesem System nicht verfügbar.");
    }
}
