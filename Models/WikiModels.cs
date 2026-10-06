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
public sealed record WikiKnowledgePage(string SourceId, string ExternalId, string Title, string Url, string SpaceKey, string Version, DateTime? LastModifiedUtc);
public sealed record WikiKnowledgePageBatch(IReadOnlyList<WikiKnowledgePage> Pages, bool HasMore);
public sealed record WikiKnowledgePageContent(string ExternalId, string Title, string PlainText, string Version, DateTime? LastModifiedUtc, string? StorageMarkup = null);
public sealed record WikiAiIndexStatus(string SourceId, int PageCount, int ChunkCount, DateTime? LastSuccessUtc, string Status, int ProcessedPages = 0, int? TotalPages = null,
    int AttachmentCount = 0, int PdfCount = 0, int PdfPageCount = 0, int ImageCount = 0, int OcrSuccessCount = 0, int OcrFailureCount = 0, int DrawIoCount = 0);

public enum WikiKnowledgeContentKind { PageText, List, Table, Code, Pdf, ImageOcr, DrawIo, AttachmentMetadata }
public sealed record WikiKnowledgeBlock(string SectionTitle, WikiKnowledgeContentKind ContentKind, string Content, int BlockOrdinal,
    string? AttachmentId = null, string? AttachmentName = null, int? PageNumber = null);
public sealed record WikiKnowledgeDocument(string SourceId, string ExternalPageId, string SpaceKey, string PageTitle, string PageUrl,
    IReadOnlyList<WikiKnowledgeBlock> Blocks);
public sealed record WikiKnowledgeAttachmentMetadata(string AttachmentId, string ParentPageId, string FileName, string MediaType,
    string Version, DateTime? LastModifiedUtc, string DownloadUrl, long? Size = null, string? AltText = null, string? Caption = null, string? SectionTitle = null);
public sealed record WikiKnowledgeAttachmentResult(WikiKnowledgeAttachmentMetadata Metadata, IReadOnlyList<WikiKnowledgeBlock> Blocks,
    string Status, string? ContentHash, int PdfPageCount = 0, bool OcrSucceeded = false, int ExtractionVersion = 1);
public sealed record WikiAttachmentPayload(byte[] Content, string Sha256);
