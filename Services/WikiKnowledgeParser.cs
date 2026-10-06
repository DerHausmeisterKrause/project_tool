using System.IO;
using System.IO.Compression;
using System.Net;
using System.Text;
using System.Xml.Linq;
using AngleSharp.Dom;
using AngleSharp.Html.Parser;
using TaskTool.Models;

namespace TaskTool.Services;

/// <summary>Turns Confluence storage HTML into stable, structure-aware text blocks.</summary>
public sealed class ConfluenceKnowledgeParser
{
    public IReadOnlySet<string> GetReferencedAttachmentNames(string? markup)
    {
        if (string.IsNullOrWhiteSpace(markup)) return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var document = new HtmlParser().ParseDocument(markup); var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var element in document.QuerySelectorAll("*"))
        {
            foreach (var attribute in element.Attributes.Where(a => a.Name.EndsWith("filename", StringComparison.OrdinalIgnoreCase)))
                if (!string.IsNullOrWhiteSpace(attribute.Value)) names.Add(attribute.Value.Trim());
            if (element.TagName.Contains("PARAMETER") && !string.IsNullOrWhiteSpace(element.TextContent))
            {
                var value = element.TextContent.Trim(); if (Path.HasExtension(value)) names.Add(value);
            }
        }
        return names;
    }

    public WikiKnowledgeDocument Parse(WikiKnowledgePage page, WikiKnowledgePageContent content)
    {
        var markup = string.IsNullOrWhiteSpace(content.StorageMarkup) ? $"<p>{WebUtility.HtmlEncode(content.PlainText)}</p>" : content.StorageMarkup;
        var document = new HtmlParser().ParseDocument(markup ?? string.Empty);
        var blocks = new List<WikiKnowledgeBlock>(); var section = string.Empty; var ordinal = 0;
        foreach (var element in document.Body?.Children.AsEnumerable() ?? Enumerable.Empty<IElement>())
        {
            var tag = element.TagName.ToLowerInvariant();
            if (tag is "h1" or "h2" or "h3" or "h4" or "h5" or "h6") { section = Clean(element.TextContent); continue; }
            if (tag == "table") Add(blocks, section, WikiKnowledgeContentKind.Table, RenderTable(element), ref ordinal);
            else if (tag is "ul" or "ol") Add(blocks, section, WikiKnowledgeContentKind.List, RenderList(element, 0), ref ordinal);
            else if (tag is "pre" or "code") Add(blocks, section, WikiKnowledgeContentKind.Code, element.TextContent.Trim(), ref ordinal);
            else if (tag is "ac:image" or "img") Add(blocks, section, WikiKnowledgeContentKind.AttachmentMetadata, RenderImage(element), ref ordinal);
            else if (tag.Contains("structured-macro") && IsDrawIo(element)) Add(blocks, section, WikiKnowledgeContentKind.AttachmentMetadata, RenderDrawIoReference(element), ref ordinal);
            else Add(blocks, section, WikiKnowledgeContentKind.PageText, Clean(element.TextContent), ref ordinal);
        }
        if (blocks.Count == 0) Add(blocks, string.Empty, WikiKnowledgeContentKind.PageText, content.PlainText, ref ordinal);
        return new(page.SourceId, page.ExternalId, page.SpaceKey, page.Title, page.Url, blocks);
    }

    private static void Add(List<WikiKnowledgeBlock> blocks, string section, WikiKnowledgeContentKind kind, string value, ref int ordinal)
    {
        value = value.Trim(); if (value.Length == 0) return;
        blocks.Add(new(section, kind, value, ordinal++));
    }
    private static string Clean(string? value) => string.Join(' ', (value ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
    private static string RenderTable(IElement table)
    {
        var rows = table.QuerySelectorAll("tr").Select(row => row.Children.Where(cell => cell.TagName.Equals("TH", StringComparison.OrdinalIgnoreCase) || cell.TagName.Equals("TD", StringComparison.OrdinalIgnoreCase)).Select(cell => Clean(cell.TextContent)).ToArray()).Where(row => row.Length > 0).ToArray();
        if (rows.Length == 0) return string.Empty;
        var width = rows.Max(row => row.Length); var header = rows[0];
        string Row(IReadOnlyList<string> row) => "| " + string.Join(" | ", Enumerable.Range(0, width).Select(i => (i < row.Count ? row[i] : string.Empty).Replace("|", "\\|"))) + " |";
        return string.Join('\n', new[] { Row(header), "| " + string.Join(" | ", Enumerable.Repeat("---", width)) + " |" }.Concat(rows.Skip(1).Select(Row)));
    }
    private static string RenderList(IElement list, int depth)
    {
        var ordered = list.TagName.Equals("OL", StringComparison.OrdinalIgnoreCase); var lines = new List<string>(); var index = 1;
        foreach (var item in list.Children.Where(x => x.TagName.Equals("LI", StringComparison.OrdinalIgnoreCase)))
        {
            var own = string.Concat(item.ChildNodes.Where(x => x is not IElement e || (e.TagName != "UL" && e.TagName != "OL")).Select(x => x.TextContent));
            lines.Add($"{new string(' ', depth * 2)}{(ordered ? $"{index++}." : "-")} {Clean(own)}");
            foreach (var nested in item.Children.Where(x => x.TagName is "UL" or "OL")) lines.Add(RenderList(nested, depth + 1));
        }
        return string.Join('\n', lines);
    }
    private static string RenderImage(IElement image)
    {
        var name = image.GetAttribute("ri:filename") ?? image.QuerySelector("ri\\:attachment")?.GetAttribute("ri:filename") ?? image.GetAttribute("src") ?? "Bild";
        var alt = image.GetAttribute("alt") ?? image.GetAttribute("ac:alt") ?? string.Empty;
        return $"[Bild: {name}]" + (alt.Length == 0 ? string.Empty : $"\nAlt-Text: {alt}");
    }
    private static bool IsDrawIo(IElement element) => (element.GetAttribute("ac:name") ?? element.GetAttribute("data-macro-name") ?? string.Empty).Contains("drawio", StringComparison.OrdinalIgnoreCase);
    private static string RenderDrawIoReference(IElement element) => "Draw.io Diagramm: " + (element.QuerySelector("ac\\:parameter")?.TextContent?.Trim() ?? "Diagramm");
}

public static class WikiKnowledgeChunker
{
    public const int MaximumCharacters = 1200;
    public static IReadOnlyList<WikiKnowledgeBlock> Chunk(WikiKnowledgeDocument document)
    {
        var result = new List<WikiKnowledgeBlock>();
        foreach (var block in document.Blocks)
        {
            if (block.Content.Length <= MaximumCharacters) { result.Add(block); continue; }
            var header = block.ContentKind == WikiKnowledgeContentKind.Table ? block.Content.Split('\n').Take(2).ToArray() : [];
            var sourceLines = block.ContentKind == WikiKnowledgeContentKind.Table ? block.Content.Split('\n').Skip(2) : SplitLongLines(block.Content);
            var body = sourceLines.SelectMany(line => line.Length <= MaximumCharacters ? [line] : Enumerable.Range(0, (line.Length + MaximumCharacters - 1) / MaximumCharacters).Select(index => line.Substring(index * MaximumCharacters, Math.Min(MaximumCharacters, line.Length - index * MaximumCharacters))));
            var current = new StringBuilder();
            foreach (var line in body)
            {
                if (current.Length > 0 && current.Length + line.Length + 1 > MaximumCharacters)
                {
                    result.Add(block with { Content = Prefix(header, current.ToString()) }); current.Clear();
                }
                current.AppendLine(line);
            }
            if (current.Length > 0) result.Add(block with { Content = Prefix(header, current.ToString()) });
        }
        return result;
    }
    private static string Prefix(string[] header, string content) => header.Length == 0 ? content.Trim() : string.Join('\n', header) + "\n" + content.Trim();
    private static IEnumerable<string> SplitLongLines(string content) => content.Split('\n');
}

public sealed class WikiDrawIoExtractor : IWikiDrawIoExtractor
{
    public string Extract(string xml)
    {
        var root = XDocument.Parse(xml); var model = root.Root?.Name.LocalName == "mxGraphModel" ? root.Root : root.Descendants("mxGraphModel").FirstOrDefault() ?? DecodeDiagram(root.Descendants("diagram").FirstOrDefault()?.Value);
        if (model == null) throw new InvalidDataException("Draw.io diagram data is invalid.");
        var cells = model.Descendants("mxCell").ToDictionary(x => (string?)x.Attribute("id") ?? Guid.NewGuid().ToString(), x => x);
        var nodes = cells.Values.Where(x => (string?)x.Attribute("vertex") == "1").Select(x => StripHtml((string?)x.Attribute("value"))).Where(x => x.Length > 0).ToArray();
        var edges = cells.Values.Where(x => (string?)x.Attribute("edge") == "1").Select(x =>
        {
            var source = Label(cells, (string?)x.Attribute("source")); var target = Label(cells, (string?)x.Attribute("target")); var label = StripHtml((string?)x.Attribute("value"));
            return source.Length == 0 || target.Length == 0 ? string.Empty : $"- {source} -> {target}" + (label.Length == 0 ? string.Empty : $" [{label}]");
        }).Where(x => x.Length > 0);
        return "Knoten:\n" + string.Join('\n', nodes.Select(x => $"- {x}")) + "\n\nVerbindungen:\n" + string.Join('\n', edges);
    }
    private static XElement? DecodeDiagram(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;
        var compressed = Convert.FromBase64String(value); using var input = new MemoryStream(compressed); using var deflate = new DeflateStream(input, CompressionMode.Decompress); using var reader = new StreamReader(deflate);
        return XDocument.Parse(Uri.UnescapeDataString(reader.ReadToEnd())).Root;
    }
    private static string Label(IReadOnlyDictionary<string, XElement> cells, string? id) => id != null && cells.TryGetValue(id, out var cell) ? StripHtml((string?)cell.Attribute("value")) : string.Empty;
    private static string StripHtml(string? value) => WebUtility.HtmlDecode(new HtmlParser().ParseDocument(value ?? string.Empty).Body?.TextContent ?? string.Empty).Trim();
}
