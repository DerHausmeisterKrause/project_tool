using System.IO;
using System.Text;
using System.Threading;
using UglyToad.PdfPig;

namespace TaskTool.Services;

public sealed record ExtractedKnowledgeText(string Text, int? PageNumber = null);

public interface IAiKnowledgeDocumentExtractor
{
    bool IsSupported(string path);
    Task<IReadOnlyList<ExtractedKnowledgeText>> ExtractAsync(string path, CancellationToken cancellationToken = default);
}

public sealed class AiKnowledgeDocumentExtractor : IAiKnowledgeDocumentExtractor
{
    private static readonly HashSet<string> Extensions = new(StringComparer.OrdinalIgnoreCase) { ".txt", ".md", ".pdf" };
    public bool IsSupported(string path) => Extensions.Contains(Path.GetExtension(path));

    public async Task<IReadOnlyList<ExtractedKnowledgeText>> ExtractAsync(string path, CancellationToken cancellationToken = default)
    {
        var extension = Path.GetExtension(path);
        if (!IsSupported(path)) return Array.Empty<ExtractedKnowledgeText>();
        if (!extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            return new[] { new ExtractedKnowledgeText(await File.ReadAllTextAsync(path, Encoding.UTF8, cancellationToken)) };

        return await Task.Run<IReadOnlyList<ExtractedKnowledgeText>>(() =>
        {
            using var document = PdfDocument.Open(path);
            return document.GetPages()
                .Select(page => new ExtractedKnowledgeText(page.Text, page.Number))
                .Where(page => !string.IsNullOrWhiteSpace(page.Text))
                .ToArray();
        }, cancellationToken);
    }
}
