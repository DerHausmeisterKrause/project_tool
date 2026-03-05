using System.Text.RegularExpressions;
using TaskTool.Models;

namespace TaskTool.ViewModels;

internal static class OutlookAllDayMarkerMapper
{
    // Mapping examples (expected):
    // - "Homeoffice von Asböck Marcel" => HO (SubjectPrefix:Homeoffice)
    // - "Workshop" (all-day) => null
    // - "Homepage Relaunch" (all-day) => null
    // - "Urlaub" => UL (SubjectPrefix:Urlaub)
    // - "MAZ Bereitschaft" => AM (SubjectPrefix:MAZ)
    // - "Team Homeoffice Planung" => null (no strict token/prefix/category/tag match)
    public static string? TryMapAllDayMarker(OutlookCalendarEvent evt, out string matchedBy)
    {
        matchedBy = "NoMatch";

        var subject = Normalize(evt.Subject);
        var categories = ParseCategories(evt.Categories);
        var body = Normalize(evt.BodyPreview);

        if (MatchesHo(subject, categories, body, out matchedBy))
            return "HO";

        if (MatchesUl(subject, categories, body, out matchedBy))
            return "UL";

        if (MatchesAm(subject, categories, body, out matchedBy))
            return "AM";

        return null;
    }

    private static bool MatchesHo(string subject, HashSet<string> categories, string body, out string matchedBy)
    {
        matchedBy = "NoMatch";
        if (StartsWithToken(subject, "HO")) { matchedBy = "SubjectPrefix:HO"; return true; }
        if (StartsWithToken(subject, "HOMEOFFICE")) { matchedBy = "SubjectPrefix:Homeoffice"; return true; }
        if (ContainsBracketToken(subject, "HO")) { matchedBy = "SubjectToken:[HO]/(HO)"; return true; }
        if (ContainsStandaloneToken(subject, "HO")) { matchedBy = "SubjectToken:HO"; return true; }
        if (categories.Contains("HOMEOFFICE")) { matchedBy = "Category:Homeoffice"; return true; }
        if (categories.Contains("HO")) { matchedBy = "Category:HO"; return true; }
        if (ContainsHashTag(body, "HO")) { matchedBy = "BodyTag:#HO"; return true; }
        return false;
    }

    private static bool MatchesUl(string subject, HashSet<string> categories, string body, out string matchedBy)
    {
        matchedBy = "NoMatch";
        if (StartsWithToken(subject, "UL")) { matchedBy = "SubjectPrefix:UL"; return true; }
        if (StartsWithToken(subject, "URLAUB")) { matchedBy = "SubjectPrefix:Urlaub"; return true; }
        if (ContainsBracketToken(subject, "UL")) { matchedBy = "SubjectToken:[UL]/(UL)"; return true; }
        if (ContainsStandaloneToken(subject, "UL")) { matchedBy = "SubjectToken:UL"; return true; }
        if (categories.Contains("URLAUB")) { matchedBy = "Category:Urlaub"; return true; }
        if (categories.Contains("UL")) { matchedBy = "Category:UL"; return true; }
        if (ContainsHashTag(body, "UL")) { matchedBy = "BodyTag:#UL"; return true; }
        return false;
    }

    private static bool MatchesAm(string subject, HashSet<string> categories, string body, out string matchedBy)
    {
        matchedBy = "NoMatch";
        if (StartsWithToken(subject, "AM")) { matchedBy = "SubjectPrefix:AM"; return true; }
        if (StartsWithToken(subject, "MAZ")) { matchedBy = "SubjectPrefix:MAZ"; return true; }
        if (ContainsBracketToken(subject, "AM")) { matchedBy = "SubjectToken:[AM]/(AM)"; return true; }
        if (ContainsStandaloneToken(subject, "MAZ")) { matchedBy = "SubjectToken:MAZ"; return true; }
        if (categories.Contains("AM")) { matchedBy = "Category:AM"; return true; }
        if (categories.Contains("MAZ")) { matchedBy = "Category:MAZ"; return true; }
        if (ContainsHashTag(body, "AM")) { matchedBy = "BodyTag:#AM"; return true; }
        if (ContainsHashTag(body, "MAZ")) { matchedBy = "BodyTag:#MAZ"; return true; }
        return false;
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

    private static bool StartsWithToken(string text, string token)
    {
        if (string.IsNullOrWhiteSpace(text) || string.IsNullOrWhiteSpace(token) || !text.StartsWith(token, StringComparison.Ordinal))
            return false;

        if (text.Length == token.Length)
            return true;

        return !char.IsLetterOrDigit(text[token.Length]);
    }

    private static bool ContainsBracketToken(string text, string token)
    {
        if (string.IsNullOrWhiteSpace(text) || string.IsNullOrWhiteSpace(token))
            return false;

        return text.Contains($"[{token}]", StringComparison.Ordinal)
            || text.Contains($"({token})", StringComparison.Ordinal);
    }

    private static bool ContainsStandaloneToken(string text, string token)
    {
        if (string.IsNullOrWhiteSpace(text) || string.IsNullOrWhiteSpace(token))
            return false;

        var pattern = $@"(?<![A-Z0-9]){Regex.Escape(token)}(?![A-Z0-9])";
        return Regex.IsMatch(text, pattern, RegexOptions.CultureInvariant);
    }

    private static bool ContainsHashTag(string text, string tag)
    {
        if (string.IsNullOrWhiteSpace(text) || string.IsNullOrWhiteSpace(tag))
            return false;

        var pattern = $@"(?<![A-Z0-9])#{Regex.Escape(tag)}(?![A-Z0-9])";
        return Regex.IsMatch(text, pattern, RegexOptions.CultureInvariant);
    }

    private static string Normalize(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            return string.Empty;

        return Regex.Replace(value.Trim().ToUpperInvariant(), @"\s+", " ");
    }
}
