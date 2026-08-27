namespace TaskTool.Models;

public static class ZnunySyncPolicy
{
    public const int MaximumAutomaticRequestsPerSync = 60;
    public const int MinimumSearchLimit = 10;
    public const int MaximumSearchLimit = 500;
    public const int DefaultSearchLimit = 100;
    public const int MinimumArticleLimit = 1;
    public const int MaximumArticleLimit = 100;
    public const int DefaultArticleLimit = 20;
    public const int MinimumCandidateIntervalMinutes = 3;
    public const int MaximumCandidateIntervalMinutes = 60;
    public const int DefaultCandidateIntervalMinutes = 5;
    public const string ArticleOrder = "DESC";

    public static int NormalizeIntervalMinutes(int configured)
        => configured <= 0 ? 15 : Math.Clamp(configured, 5, 1440);

    public static int NormalizeCandidateIntervalMinutes(int configured)
        => configured <= 0
            ? DefaultCandidateIntervalMinutes
            : Math.Clamp(configured, MinimumCandidateIntervalMinutes, MaximumCandidateIntervalMinutes);

    public static int NormalizeSearchLimit(int configured)
        => Math.Clamp(configured, MinimumSearchLimit, MaximumSearchLimit);

    public static int NormalizeArticleLimit(int configured)
        => Math.Clamp(configured, MinimumArticleLimit, MaximumArticleLimit);

    public static IReadOnlyList<string> SelectTicketIds(
        IEnumerable<string> currentlyAssigned,
        IEnumerable<string> previouslyAssigned)
    {
        var current = currentlyAssigned.ToHashSet(StringComparer.OrdinalIgnoreCase);
        return current.Concat(previouslyAssigned.Except(current, StringComparer.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    public static IReadOnlyDictionary<string, string> TicketGetOptions(bool allArticles, bool dynamicFields, int configuredArticleLimit)
    {
        var options = new Dictionary<string, string>
        {
            ["AllArticles"] = allArticles ? "1" : "0",
            ["DynamicFields"] = dynamicFields ? "1" : "0"
        };
        if (allArticles)
        {
            options["ArticleOrder"] = ArticleOrder;
            options["ArticleLimit"] = NormalizeArticleLimit(configuredArticleLimit).ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        return options;
    }

    public static void ApplyTicketSearchLimit(IDictionary<string, object?> payload, int configuredLimit)
        => payload["Limit"] = NormalizeSearchLimit(configuredLimit);
}
