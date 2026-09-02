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
    public static readonly TimeSpan CandidateReevaluationTtl = TimeSpan.FromMinutes(30);
    public static readonly TimeSpan DynamicFieldOptionsTtl = TimeSpan.FromHours(24);
    public const string ArticleOrder = "DESC";
    public const string SearchSortBy = "Changed";
    public const string SearchOrderBy = "Down";
    public const string AssignedOpenStateType = "Open";
    public const string AssignedNewStateType = "New";

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

    public static IReadOnlyList<string> RotateTicketIds(IEnumerable<string> ticketIds, string? nextTicketId)
    {
        var ids = ticketIds.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        if (ids.Count == 0 || string.IsNullOrWhiteSpace(nextTicketId)) return ids;
        var index = ids.FindIndex(id => string.Equals(id, nextTicketId, StringComparison.OrdinalIgnoreCase));
        return index <= 0 ? ids : ids.Skip(index).Concat(ids.Take(index)).ToList();
    }

    public static bool RequiresFullTicketGet(TicketDetailCacheEntry? cache, int requestedArticleLimit)
        => cache?.IsCompleteFor(NormalizeArticleLimit(requestedArticleLimit)) != true;

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

    public static void ApplyTicketSearchLimit(IDictionary<string, object?> payload, int configuredLimit, bool includeSorting = true)
    {
        payload["Limit"] = NormalizeSearchLimit(configuredLimit);
        if (includeSorting)
        {
            payload["SortBy"] = SearchSortBy;
            payload["OrderBy"] = SearchOrderBy;
        }
    }

    public static void ApplyTicketRoleCriteria(IDictionary<string, object?> payload, string role, int userId, bool onlyOpen)
    {
        payload[string.Equals(role, "Owner", StringComparison.Ordinal) ? "OwnerIDs" : "ResponsibleIDs"] = new[] { userId };
        if (onlyOpen) payload["StateType"] = new[] { "new", "open" };
    }

    /// <summary>
    /// Applies the scalar query contract used by the deployed GET /Ticket GenericInterface operation.
    /// This is intentionally separate from the array-based POST/candidate contract.
    /// </summary>
    public static void ApplyAssignedGetSearchCriteria(IDictionary<string, object?> payload, string role, int userId, string? stateType, int configuredLimit)
    {
        payload[string.Equals(role, "Owner", StringComparison.Ordinal) ? "OwnerIDs" : "ResponsibleIDs"] = userId;
        if (!string.IsNullOrWhiteSpace(stateType)) payload["StateType"] = stateType;
        ApplyTicketSearchLimit(payload, configuredLimit, includeSorting: false);
    }

    public static List<string> MergeTicketIds(params IEnumerable<string>[] searches)
        => searches.SelectMany(ids => ids).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
}
