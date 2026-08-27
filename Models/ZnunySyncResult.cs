namespace TaskTool.Models;

public sealed record ZnunySyncResult(
    bool Started,
    bool Success,
    bool Busy,
    int SearchResultCount,
    int UniqueTicketCount,
    int Created,
    int Updated,
    int Unchanged,
    int Skipped,
    bool SearchLimitReached,
    string ErrorMessage)
{
    public static ZnunySyncResult BusyResult() => new(false, false, true, 0, 0, 0, 0, 0, 0, false,
        "Synchronisierung konnte nicht gestartet werden, da bereits eine andere Znuny-Aktion läuft.");

    public static ZnunySyncResult Failed(string error, bool started = true) =>
        new(started, false, false, 0, 0, 0, 0, 0, 0, false, error);
}
