namespace TaskTool.Models;

public static class ZnunySyncPolicy
{
    public const int MaximumAutomaticRequestsPerSync = 60;

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
}
