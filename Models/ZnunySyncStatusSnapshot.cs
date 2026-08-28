namespace TaskTool.Models;

/// <summary>Passive, locally persisted observation of the most recent assigned-ticket full sync.</summary>
public sealed record ZnunySyncStatusSnapshot(
    DateTime TimestampUtc,
    ZnunyRequestReason Reason,
    bool Started,
    bool Success,
    bool Busy,
    int UniqueTicketCount,
    int Created,
    int Updated,
    int Unchanged,
    int Skipped,
    bool SearchLimitReached,
    string ErrorMessage);
