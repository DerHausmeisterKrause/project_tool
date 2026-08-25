using System.Security.Cryptography;
using System.Text;
using TaskTool.Models;

namespace TaskTool.Services;

public static class WikiScopePolicy
{
    public static IReadOnlyList<string> GetSpaceKeys(WikiSourceSettings source) => source.SearchAllSpaces
        ? Array.Empty<string>()
        : source.SpaceKeys.Where(x => !string.IsNullOrWhiteSpace(x)).Select(x => x.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToArray();

    public static string BuildConfluenceClause(WikiSourceSettings source)
    {
        var spaces = GetSpaceKeys(source);
        if (!source.SearchAllSpaces && spaces.Count == 0) throw new InvalidOperationException("Für die eingeschränkte Suche wurden keine Space Keys konfiguriert.");
        return spaces.Count == 0 ? string.Empty : " AND (" + string.Join(" OR ", spaces.Select(x => $"space='{ConfluenceDataCenterWikiProvider.EscapeCql(x)}'")) + ")";
    }

    public static string Fingerprint(WikiSourceSettings source)
    {
        var canonical = string.Join("|", source.ProviderType.Trim().ToLowerInvariant(), source.BaseUrl.Trim().TrimEnd('/').ToLowerInvariant(), source.SearchAllSpaces, string.Join(",", GetSpaceKeys(source)));
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(canonical))).ToLowerInvariant();
    }

    public static bool SupportsSecureVocabulary(WikiSourceSettings source) => source.ProviderType is "ConfluenceDataCenter" or "ConfluenceCloud";
}
