namespace TaskTool.Models;

public sealed record TicketDetailFetchProfile(
    bool MetadataComplete,
    bool ArticlesComplete,
    bool DynamicFieldsComplete,
    int FetchedArticleLimit,
    bool AssignmentMetadataComplete)
{
    public static TicketDetailFetchProfile Full(int articleLimit) => new(true, true, true, articleLimit, true);
    public static TicketDetailFetchProfile Candidate(int articleLimit) => new(true, true, false, articleLimit, false);
}

public sealed record TicketDetailCacheEntry(
    TicketBookingContext Context,
    string State,
    DateTime? RemoteChangedUtc,
    DateTime LastFetchedUtc,
    bool MetadataComplete,
    bool ArticlesComplete,
    bool DynamicFieldsComplete,
    int FetchedArticleLimit,
    bool AssignmentMetadataComplete)
{
    public bool IsCompleteFor(int requestedArticleLimit) => MetadataComplete
        && ArticlesComplete
        && DynamicFieldsComplete
        && FetchedArticleLimit >= requestedArticleLimit;
}

public sealed record DynamicFieldOptionsCacheEntry(
    IReadOnlyList<TicketFieldOption> Options,
    DateTime? FetchedUtc)
{
    public bool IsFresh(TimeSpan ttl, DateTime utcNow) => Options.Count > 0
        && FetchedUtc is { } fetched
        && utcNow - fetched <= ttl;
}

public sealed record TicketReconciliationCycle(
    string ContextKey,
    string DiscoveryFingerprint,
    IReadOnlyList<string> DiscoveredTicketIds,
    IReadOnlyList<string> PendingTicketIds,
    DateTime StartedUtc);
