namespace TaskTool.Models;

public enum ZnunyReconciliationWorkKind
{
    Assigned,
    RemovalVerification
}

public sealed record ZnunyReconciliationWorkItem(string TicketId, ZnunyReconciliationWorkKind WorkKind)
{
    public string PersistedKey => $"{(WorkKind == ZnunyReconciliationWorkKind.Assigned ? "assigned" : "removal")}:{TicketId}";

    public static ZnunyReconciliationWorkItem FromPersistedKey(string key)
    {
        const string assignedPrefix = "assigned:";
        const string removalPrefix = "removal:";
        if (key.StartsWith(assignedPrefix, StringComparison.OrdinalIgnoreCase))
            return new(key[assignedPrefix.Length..], ZnunyReconciliationWorkKind.Assigned);
        if (key.StartsWith(removalPrefix, StringComparison.OrdinalIgnoreCase))
            return new(key[removalPrefix.Length..], ZnunyReconciliationWorkKind.RemovalVerification);
        throw new FormatException($"Unknown Znuny reconciliation work-item key '{key}'.");
    }
}
