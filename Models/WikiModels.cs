namespace TaskTool.Models;

public sealed class WikiSearchResult
{
    public string SourceId { get; set; } = string.Empty;
    public string SourceName { get; set; } = string.Empty;
    public string ExternalId { get; set; } = string.Empty;
    public string Title { get; set; } = string.Empty;
    public string Url { get; set; } = string.Empty;
    public string Excerpt { get; set; } = string.Empty;
    public double RelevanceScore { get; set; }
    public string MatchedTerms { get; set; } = string.Empty;
    public int ProviderRank { get; set; }
    public DateTime? LastModifiedUtc { get; set; }
    public DateTime SearchedAtUtc { get; set; }
    public string RelevanceText => $"{Math.Round(RelevanceScore):0} %";
}

public sealed record WikiProviderResult(string ExternalId, string Title, string Url, string Excerpt, int ProviderRank, DateTime? LastModifiedUtc = null);
public sealed record WikiSearchSummary(int UpdatedSources, int FailedSources);
public sealed record WikiSearchTerm(string Text, string NormalizedText, double Score, string Origin, bool IsPhrase, bool WikiVocabularyMatch = false, double WikiTitleSimilarity = 0, double IdfScore = 0);
public sealed record WikiVocabularyPage(string ExternalId, string Title, string Url, string SpaceKey, DateTime? LastModifiedUtc = null);
public sealed record WikiVocabularyPageBatch(IReadOnlyList<WikiVocabularyPage> Pages, bool HasMore);
public sealed record WikiVocabularyStatus(int PageCount, DateTime? UpdatedUtc, string Status);
