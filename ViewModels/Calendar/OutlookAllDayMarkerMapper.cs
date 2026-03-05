using System.Text.RegularExpressions;
using TaskTool.Models;

namespace TaskTool.ViewModels;

internal static class OutlookAllDayMarkerMapper
{
    public static string? TryMapAllDayMarker(OutlookCalendarEvent evt, out string matchedRule)
    {
        var evaluation = Evaluate(evt);
        matchedRule = evaluation.MatchedRule;
        return evaluation.Marker;
    }

    public static string GetMatchedRule(OutlookCalendarEvent evt)
    {
        return Evaluate(evt).MatchedRule;
    }

    private static (string? Marker, string MatchedRule) Evaluate(OutlookCalendarEvent evt)
    {
        if (!evt.IsAllDay)
            return (null, "MatchedRule=Rejected(NotAllDay)");

        var subject = Normalize(evt.Subject);
        var categories = ParseCategories(evt.Categories);

        if (StartsWithMarkerPrefix(subject, "HO"))
            return ("HO", "MatchedRule=SubjectStartsWith('HO')");
        if (StartsWithMarkerPrefix(subject, "HOMEOFFICE"))
            return ("HO", "MatchedRule=SubjectStartsWith('Homeoffice')");
        if (categories.Contains("HO"))
            return ("HO", "MatchedRule=CategoryEquals('HO')");
        if (categories.Contains("HOMEOFFICE"))
            return ("HO", "MatchedRule=CategoryEquals('Homeoffice')");

        if (StartsWithMarkerPrefix(subject, "UL"))
            return ("UL", "MatchedRule=SubjectStartsWith('UL')");
        if (StartsWithMarkerPrefix(subject, "URLAUB"))
            return ("UL", "MatchedRule=SubjectStartsWith('Urlaub')");
        if (categories.Contains("UL"))
            return ("UL", "MatchedRule=CategoryEquals('UL')");
        if (categories.Contains("URLAUB"))
            return ("UL", "MatchedRule=CategoryEquals('Urlaub')");

        if (StartsWithMarkerPrefix(subject, "AM"))
            return ("AM", "MatchedRule=SubjectStartsWith('AM')");
        if (StartsWithMarkerPrefix(subject, "MAZ"))
            return ("AM", "MatchedRule=SubjectStartsWith('MAZ')");
        if (categories.Contains("AM"))
            return ("AM", "MatchedRule=CategoryEquals('AM')");
        if (categories.Contains("MAZ"))
            return ("AM", "MatchedRule=CategoryEquals('MAZ')");

        return (null, "MatchedRule=NoMatch(StrictAllDayPrefixOrCategory)");
    }

    private static bool StartsWithMarkerPrefix(string text, string prefix)
    {
        if (string.IsNullOrWhiteSpace(text) || string.IsNullOrWhiteSpace(prefix))
            return false;

        if (!text.StartsWith(prefix, StringComparison.Ordinal))
            return false;

        if (text.Length == prefix.Length)
            return true;

        return !char.IsLetterOrDigit(text[prefix.Length]);
    }

    private static HashSet<string> ParseCategories(string categories)
    {
        if (string.IsNullOrWhiteSpace(categories))
            return new HashSet<string>(StringComparer.Ordinal);

        return categories
            .Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(c => Normalize(c))
            .Where(c => !string.IsNullOrWhiteSpace(c))
            .ToHashSet(StringComparer.Ordinal);
    }

    private static string Normalize(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        return Regex.Replace(value.Trim().ToUpperInvariant(), @"\s+", " ");
    }
}
