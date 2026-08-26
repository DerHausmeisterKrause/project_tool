namespace TaskTool.Models;

public static class TicketPendingState
{
    public static bool IsPendingStateType(string? stateType)
        => string.Equals(stateType?.Trim(), "pending reminder", StringComparison.OrdinalIgnoreCase)
           || string.Equals(stateType?.Trim(), "pending auto", StringComparison.OrdinalIgnoreCase);

    public static bool IsActive(TaskItem task, DateTime utcNow)
        => task.IsActivelyTicketPending(utcNow);

    public static bool IsWakeCandidate(TaskItem task, DateTime utcNow)
        => task.IsZnunyTask && IsPendingStateType(task.TicketStateType)
           && task.TicketPendingUntilUtc is DateTime until
           && NormalizePendingUtc(until) >= new DateTime(2000, 1, 1, 0, 0, 0, DateTimeKind.Utc)
           && NormalizePendingUtc(until) <= NormalizePendingUtc(utcNow)
           && !WasHandledFor(task, until);

    public static DateTime? ResolveRelativePendingUtc(DateTime responseReceivedUtc, long untilTimeSeconds)
    {
        try
        {
            return NormalizePendingUtc(responseReceivedUtc).AddSeconds(untilTimeSeconds);
        }
        catch (ArgumentOutOfRangeException)
        {
            return null;
        }
    }

    public static DateTime NormalizePendingUtc(DateTime value)
    {
        var utc = value.Kind switch
        {
            DateTimeKind.Utc => value,
            DateTimeKind.Local => value.ToUniversalTime(),
            _ => DateTime.SpecifyKind(value, DateTimeKind.Utc)
        };
        return new DateTime(utc.Year, utc.Month, utc.Day, utc.Hour, utc.Minute, utc.Second, DateTimeKind.Utc);
    }

    public static long ToUnixSeconds(DateTime value)
        => new DateTimeOffset(NormalizePendingUtc(value)).ToUnixTimeSeconds();

    public static bool IsSameWake(DateTime? left, DateTime? right)
        => left.HasValue && right.HasValue && ToUnixSeconds(left.Value) == ToUnixSeconds(right.Value);

    public static string CreateWakeKey(string ticketId, DateTime pendingUntilUtc)
        => $"{ticketId.Trim()}|{ToUnixSeconds(pendingUntilUtc)}";

    public static bool WasHandledFor(TaskItem task, DateTime pendingUntilUtc)
        => IsSameWake(task.PendingWakeHandledForUtc, pendingUntilUtc);

    public static bool WasNotificationSentFor(TaskItem task, DateTime pendingUntilUtc)
        => IsSameWake(task.PendingWakeNotificationForUtc, pendingUntilUtc);
}
