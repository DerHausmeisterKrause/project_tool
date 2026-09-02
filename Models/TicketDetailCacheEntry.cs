namespace TaskTool.Models;

public sealed record TicketDetailFetchProfile(
    bool MetadataComplete,
    bool ArticlesComplete,
    bool DynamicFieldsComplete,
    int FetchedArticleLimit)
{
    public static TicketDetailFetchProfile Full(int articleLimit) => new(true, true, true, articleLimit);
    public static TicketDetailFetchProfile Candidate(int articleLimit) => new(true, true, false, articleLimit);
}

public sealed record TicketDetailCacheEntry(
    TicketBookingContext Context,
    string State,
    DateTime? RemoteChangedUtc,
    DateTime LastFetchedUtc,
    bool MetadataComplete,
    bool ArticlesComplete,
    bool DynamicFieldsComplete,
    int FetchedArticleLimit)
{
    public bool IsCompleteFor(int requestedArticleLimit) => MetadataComplete
        && ArticlesComplete
        && DynamicFieldsComplete
        && FetchedArticleLimit >= requestedArticleLimit;
}
