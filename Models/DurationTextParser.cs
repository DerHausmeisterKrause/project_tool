using System.Globalization;

namespace TaskTool.Models;

public static class DurationTextParser
{
    public static bool TryParseSeconds(string? text, out long seconds)
    {
        seconds = 0;
        var parts = (text ?? string.Empty).Trim().Split(':');
        if (parts.Length != 3 || parts.Any(part => part.Length == 0 || !long.TryParse(part, NumberStyles.None, CultureInfo.InvariantCulture, out _)))
            return false;
        if (!long.TryParse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture, out var hours)
            || !int.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out var minutes)
            || !int.TryParse(parts[2], NumberStyles.None, CultureInfo.InvariantCulture, out var remainingSeconds)
            || minutes is < 0 or > 59 || remainingSeconds is < 0 or > 59)
            return false;
        try { seconds = checked(hours * 3600 + minutes * 60L + remainingSeconds); return true; }
        catch (OverflowException) { return false; }
    }

    public static string Format(long seconds)
    {
        seconds = Math.Max(0, seconds);
        return $"{seconds / 3600:00}:{seconds % 3600 / 60:00}:{seconds % 60:00}";
    }
}
