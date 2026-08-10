namespace TaskTool.Services;

public sealed class GermanTimeService
{
    private static readonly string[] TimeZoneIds = ["W. Europe Standard Time", "Europe/Berlin"];

    public TimeZoneInfo TimeZone { get; } = ResolveTimeZone();

    public DateTime GetGermanLocalNow(DateTime? utcNow = null)
    {
        var utc = utcNow ?? DateTime.UtcNow;
        if (utc.Kind != DateTimeKind.Utc)
            utc = utc.ToUniversalTime();

        return TimeZoneInfo.ConvertTimeFromUtc(utc, TimeZone);
    }

    private static TimeZoneInfo ResolveTimeZone()
    {
        foreach (var id in TimeZoneIds)
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

        throw new TimeZoneNotFoundException("Weder 'W. Europe Standard Time' noch 'Europe/Berlin' ist auf diesem System verfügbar.");
    }
}
