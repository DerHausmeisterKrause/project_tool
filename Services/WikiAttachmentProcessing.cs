using System.IO;
using System.Security.Cryptography;
using System.Text;
using TaskTool.Models;
using UglyToad.PdfPig;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage.Streams;
using WindowsPdfDocument = Windows.Data.Pdf.PdfDocument;

namespace TaskTool.Services;

public interface IWikiPdfExtractor { Task<IReadOnlyList<ExtractedKnowledgeText>> ExtractAsync(byte[] content, CancellationToken token); }
public interface IWikiImageOcrService { bool IsAvailable { get; } Task<string> RecognizeAsync(byte[] content, string mediaType, CancellationToken token); }
public interface IWikiDrawIoExtractor { string Extract(string xml); }

public sealed class WikiPdfExtractor(IWikiImageOcrService? ocr = null) : IWikiPdfExtractor
{
    public async Task<IReadOnlyList<ExtractedKnowledgeText>> ExtractAsync(byte[] content, CancellationToken token)
    {
        var pages=await Task.Run(()=>{using var document=PdfDocument.Open(content);return document.GetPages().Select(page=>new ExtractedKnowledgeText(page.Text,page.Number)).ToArray();},token);
        if(ocr?.IsAvailable!=true||pages.All(page=>HasMeaningfulText(page.Text)))return pages.Where(page=>!string.IsNullOrWhiteSpace(page.Text)).ToArray();
        using var input=new InMemoryRandomAccessStream();using(var writer=new DataWriter(input)){writer.WriteBytes(content);await writer.StoreAsync();writer.DetachStream();}input.Seek(0);var rendered=await WindowsPdfDocument.LoadFromStreamAsync(input);var result=new List<ExtractedKnowledgeText>();
        foreach(var page in pages){token.ThrowIfCancellationRequested();if(HasMeaningfulText(page.Text)){result.Add(page);continue;}using var pdfPage=rendered.GetPage((uint)(page.PageNumber!.Value-1));using var image=new InMemoryRandomAccessStream();await pdfPage.RenderToStreamAsync(image);image.Seek(0);using var reader=new DataReader(image.GetInputStreamAt(0));await reader.LoadAsync((uint)image.Size);var bytes=new byte[checked((int)image.Size)];reader.ReadBytes(bytes);var text=await ocr.RecognizeAsync(bytes,"image/png",token);if(!string.IsNullOrWhiteSpace(text))result.Add(new(text,page.PageNumber));}
        return result;
    }
    private static bool HasMeaningfulText(string? text)=>!string.IsNullOrWhiteSpace(text)&&text.Count(char.IsLetterOrDigit)>=20;
}

/// <summary>Safe default when the optional Windows OCR runtime/language pack is unavailable.</summary>
public sealed class WindowsWikiImageOcrService : IWikiImageOcrService
{
    private static readonly SemaphoreSlim Throttle=new(2,2);
    public WindowsWikiImageOcrService(string? localAppData = null) { }
    public bool IsAvailable => OperatingSystem.IsWindowsVersionAtLeast(10, 0, 17763) && OcrEngine.TryCreateFromUserProfileLanguages() != null;
    public async Task<string> RecognizeAsync(byte[] content, string mediaType, CancellationToken token)
    {
        if (!IsAvailable) return string.Empty;
        await Throttle.WaitAsync(token);try{token.ThrowIfCancellationRequested(); using var stream = new InMemoryRandomAccessStream(); using (var writer = new DataWriter(stream)) { writer.WriteBytes(content); await writer.StoreAsync(); writer.DetachStream(); }
        stream.Seek(0); var decoder = await BitmapDecoder.CreateAsync(stream); using var bitmap = await decoder.GetSoftwareBitmapAsync(); token.ThrowIfCancellationRequested();
        var engine = OcrEngine.TryCreateFromUserProfileLanguages(); if (engine == null) return string.Empty; var result = await engine.RecognizeAsync(bitmap); token.ThrowIfCancellationRequested(); return NormalizeOcrText(result.Text);}finally{Throttle.Release();}
    }
    private static string NormalizeOcrText(string? value) => string.Join('\n', (value ?? string.Empty).Normalize().Replace("\0", string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n').Select(line => string.Join(' ', line.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries))).Where(line => line.Length > 0));
}

public sealed class WikiAttachmentProcessor
{
    public const int CurrentExtractionVersion = 1;
    private static readonly SemaphoreSlim OcrThrottle = new(2, 2);
    public const long MaximumPdfBytes = 50L * 1024 * 1024;
    public const long MaximumImageBytes = 20L * 1024 * 1024;
    public const long MaximumDrawIoBytes = 20L * 1024 * 1024;
    private static readonly HashSet<string> ImageExtensions = new(StringComparer.OrdinalIgnoreCase) { ".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff" };
    private readonly IWikiPdfExtractor _pdf; private readonly IWikiImageOcrService _ocr; private readonly IWikiDrawIoExtractor _drawIo;
    public WikiAttachmentProcessor(IWikiPdfExtractor? pdf = null, IWikiImageOcrService? ocr = null, IWikiDrawIoExtractor? drawIo = null)
    { _ocr = ocr ?? new WindowsWikiImageOcrService(); _pdf = pdf ?? new WikiPdfExtractor(_ocr); _drawIo = drawIo ?? new WikiDrawIoExtractor(); }

    public async Task<WikiKnowledgeAttachmentResult> ProcessAsync(WikiKnowledgePage page, WikiKnowledgeAttachmentMetadata metadata, Stream stream, CancellationToken token)
    {
        var extension = Path.GetExtension(metadata.FileName); var isPdf = extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase) || metadata.MediaType.Equals("application/pdf", StringComparison.OrdinalIgnoreCase);
        var isImage = ImageExtensions.Contains(extension) || metadata.MediaType.StartsWith("image/", StringComparison.OrdinalIgnoreCase);
        var isDraw = extension.Equals(".drawio", StringComparison.OrdinalIgnoreCase) || extension.Equals(".xml", StringComparison.OrdinalIgnoreCase) || metadata.MediaType.Contains("drawio", StringComparison.OrdinalIgnoreCase);
        var limit = isPdf ? MaximumPdfBytes : isImage ? MaximumImageBytes : isDraw ? MaximumDrawIoBytes : 0;
        if (limit == 0) return new(metadata, [MetadataBlock(metadata)], "unsupported", null);
        if (metadata.Size > limit) return new(metadata, [MetadataBlock(metadata)], "skipped-too-large", null);
        var payload = await ReadPayloadAsync(stream, limit, token); return await ProcessPayloadAsync(metadata, payload, isPdf, isImage, isDraw, token);
    }
    public async Task<WikiAttachmentPayload> ReadPayloadAsync(WikiKnowledgeAttachmentMetadata metadata, Stream stream, CancellationToken token)
    {
        var extension=Path.GetExtension(metadata.FileName);var limit=extension.Equals(".pdf",StringComparison.OrdinalIgnoreCase)||metadata.MediaType.Equals("application/pdf",StringComparison.OrdinalIgnoreCase)?MaximumPdfBytes:ImageExtensions.Contains(extension)||metadata.MediaType.StartsWith("image/",StringComparison.OrdinalIgnoreCase)?MaximumImageBytes:MaximumDrawIoBytes;
        return await ReadPayloadAsync(stream,limit,token);
    }
    public async Task<WikiKnowledgeAttachmentResult> ProcessPayloadAsync(WikiKnowledgeAttachmentMetadata metadata, WikiAttachmentPayload payload, CancellationToken token)
    {
        var extension=Path.GetExtension(metadata.FileName);var isPdf=extension.Equals(".pdf",StringComparison.OrdinalIgnoreCase)||metadata.MediaType.Equals("application/pdf",StringComparison.OrdinalIgnoreCase);var isImage=ImageExtensions.Contains(extension)||metadata.MediaType.StartsWith("image/",StringComparison.OrdinalIgnoreCase);var isDraw=extension.Equals(".drawio",StringComparison.OrdinalIgnoreCase)||extension.Equals(".xml",StringComparison.OrdinalIgnoreCase)||metadata.MediaType.Contains("drawio",StringComparison.OrdinalIgnoreCase);
        return await ProcessPayloadAsync(metadata,payload,isPdf,isImage,isDraw,token);
    }
    private async Task<WikiKnowledgeAttachmentResult> ProcessPayloadAsync(WikiKnowledgeAttachmentMetadata metadata,WikiAttachmentPayload payload,bool isPdf,bool isImage,bool isDraw,CancellationToken token)
    {
        var bytes=payload.Content;var hash=payload.Sha256;
        if (isPdf)
        {
            var pages = await _pdf.ExtractAsync(bytes, token); var blocks = pages.Select((text, index) => new WikiKnowledgeBlock(metadata.SectionTitle ?? string.Empty, WikiKnowledgeContentKind.Pdf, text.Text, index, metadata.AttachmentId, metadata.FileName, text.PageNumber)).ToArray();
            return new(metadata, blocks.Length == 0 ? [MetadataBlock(metadata)] : blocks, blocks.Length == 0 ? "metadata-only" : "indexed", hash, pages.Count);
        }
        if (isDraw)
        {
            var text = _drawIo.Extract(Encoding.UTF8.GetString(bytes));
            return new(metadata, [new(metadata.SectionTitle ?? string.Empty, WikiKnowledgeContentKind.DrawIo, $"Draw.io Diagramm: {metadata.FileName}\n\n{text}", 0, metadata.AttachmentId, metadata.FileName)], "indexed", hash);
        }
        var prefix = $"[Bild: {metadata.FileName}]" + (string.IsNullOrWhiteSpace(metadata.AltText) ? string.Empty : $"\nAlt-Text: {metadata.AltText}") + (string.IsNullOrWhiteSpace(metadata.Caption) ? string.Empty : $"\nCaption: {metadata.Caption}");
        if (!_ocr.IsAvailable) return new(metadata, [new(metadata.SectionTitle ?? string.Empty, WikiKnowledgeContentKind.AttachmentMetadata, prefix, 0, metadata.AttachmentId, metadata.FileName)], "ocr-unavailable", hash);
        try
        {
            await OcrThrottle.WaitAsync(token); string ocr; try { ocr = await _ocr.RecognizeAsync(bytes, metadata.MediaType, token); } finally { OcrThrottle.Release(); }
            var meaningful = ocr.Split((char[]?)null,StringSplitOptions.RemoveEmptyEntries).Sum(x=>x.Length) >= 8;
            return new(metadata, [new(metadata.SectionTitle ?? string.Empty, meaningful ? WikiKnowledgeContentKind.ImageOcr : WikiKnowledgeContentKind.AttachmentMetadata, prefix + (meaningful ? $"\nOCR:\n{ocr}" : string.Empty), 0, metadata.AttachmentId, metadata.FileName)], meaningful ? "ocr-success" : "ocr-empty", hash, OcrSucceeded: meaningful);
        }
        catch (OperationCanceledException) { throw; }
        catch { return new(metadata,[new(metadata.SectionTitle ?? string.Empty,WikiKnowledgeContentKind.AttachmentMetadata,prefix,0,metadata.AttachmentId,metadata.FileName)],"ocr-failed",hash); }
    }

    private static WikiKnowledgeBlock MetadataBlock(WikiKnowledgeAttachmentMetadata value) => new(value.SectionTitle ?? string.Empty, WikiKnowledgeContentKind.AttachmentMetadata, $"Attachment: {value.FileName}", 0, value.AttachmentId, value.FileName);
    private static async Task<WikiAttachmentPayload> ReadPayloadAsync(Stream stream, long maximum, CancellationToken token)
    {
        using var output = new MemoryStream(); var buffer = new byte[81920]; long total = 0;
        for (var read = await stream.ReadAsync(buffer, token); read > 0; read = await stream.ReadAsync(buffer, token)) { total += read; if (total > maximum) throw new InvalidDataException("Attachment exceeds the configured size limit."); await output.WriteAsync(buffer.AsMemory(0, read), token); }
        var bytes=output.ToArray();return new(bytes,Convert.ToHexString(SHA256.HashData(bytes)));
    }
}
