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
           && task.TicketPendingUntilUtc is DateTime until && until <= utcNow;

    public static bool WasNotificationSentFor(TaskItem task, DateTime pendingUntilUtc)
        => task.PendingWakeNotificationForUtc is DateTime sent
           && sent.ToUniversalTime() == pendingUntilUtc.ToUniversalTime();
}
