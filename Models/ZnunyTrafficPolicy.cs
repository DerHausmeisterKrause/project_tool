namespace TaskTool.Models;

public static class ZnunyTrafficPolicy
{
    public const int MinimumSyncIntervalMinutes = 5;
    public const int AutomaticRequestsPerMinute = 30;
    public const int CandidateTicketDetailsPerRefresh = 25;

    public static int NormalizeSyncIntervalMinutes(int configuredMinutes)
        => configuredMinutes <= 0 ? 15 : Math.Clamp(configuredMinutes, MinimumSyncIntervalMinutes, 1440);

    public static IReadOnlyList<string> SelectAssignedSyncTicketIds(
        IEnumerable<string> currentlyAssignedIds,
        IEnumerable<string> previouslyAssignedIds)
    {
        var current = currentlyAssignedIds.ToHashSet(StringComparer.OrdinalIgnoreCase);
        return current
            .Concat(previouslyAssignedIds.Except(current, StringComparer.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }
}
