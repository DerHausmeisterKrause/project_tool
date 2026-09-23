namespace TaskTool.Services;

public sealed record KnowledgeChunkText(string Content, int? PageNumber);

public static class AiKnowledgeChunker
{
    public const int TargetSize = 1000;
    public const int MaximumSize = 1200;
    public const int Overlap = 180;

    public static IReadOnlyList<KnowledgeChunkText> Chunk(IEnumerable<ExtractedKnowledgeText> sections)
    {
        var result = new List<KnowledgeChunkText>();
        foreach (var section in sections)
        {
            var text = section.Text.Replace("\r\n", "\n").Trim();
            var start = 0;
            while (start < text.Length)
            {
                var end = Math.Min(start + MaximumSize, text.Length);
                if (end < text.Length)
                {
                    var preferred = text.LastIndexOf('\n', end - 1, end - start);
                    if (preferred < start + TargetSize / 2) preferred = text.LastIndexOf(' ', end - 1, end - start);
                    if (preferred >= start + TargetSize / 2) end = preferred + 1;
                }
                var content = text[start..end].Trim();
                if (content.Length > 0) result.Add(new(content, section.PageNumber));
                if (end >= text.Length) break;
                start = Math.Max(start + 1, end - Overlap);
            }
        }
        return result;
    }
}
