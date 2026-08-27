namespace TaskTool.Models;

public static class ZnunySyncPolicy
{
    public const int MaximumAutomaticRequestsPerSync = 60;
    public const int TicketSearchLimit = 100;
    public const int DefaultArticleLimit = 20;
    public const string ArticleOrder = "DESC";

    public static int NormalizeIntervalMinutes(int configured)
        => configured <= 0 ? 15 : Math.Clamp(configured, 5, 1440);

    public static IReadOnlyList<string> SelectTicketIds(
        IEnumerable<string> currentlyAssigned,
        IEnumerable<string> previouslyAssigned)
    {
        var current = currentlyAssigned.ToHashSet(StringComparer.OrdinalIgnoreCase);
        return current.Concat(previouslyAssigned.Except(current, StringComparer.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    public static IReadOnlyDictionary<string, string> TicketGetOptions(bool allArticles, bool dynamicFields)
    {
        var options = new Dictionary<string, string>
        {
            ["AllArticles"] = allArticles ? "1" : "0",
            ["DynamicFields"] = dynamicFields ? "1" : "0"
        };
        if (allArticles)
        {
            options["ArticleOrder"] = ArticleOrder;
            options["ArticleLimit"] = DefaultArticleLimit.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
        return options;
    }

    public static void ApplyTicketSearchLimit(IDictionary<string, object?> payload)
        => payload["Limit"] = TicketSearchLimit;
}
