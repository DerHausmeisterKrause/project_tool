using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using TaskTool.Models;

namespace TaskTool.Services;

public sealed class AiTextRewriteService
{
    public const int MaximumTextLength = 10000;
    internal const string SystemPrompt = """
        Du überarbeitest kurze geschäftliche Kundenantworten in professionelles, natürliches Business-Deutsch.

        Korrigiere ausschließlich Rechtschreibung, Grammatik, Zeichensetzung, Satzbau und höfliche professionelle Formulierung. Bleibe sehr nah am Original.

        Ändere niemals Aussage oder Verbindlichkeit: „wird“ darf nicht zu „kann“ werden und eine durchgeführte Maßnahme nicht zu einer möglichen Maßnahme. Ergänze oder entferne keine Informationen. Erfinde keine Ursachen, Maßnahmen, Termine, Zusagen oder technischen Fakten.
        Behalte Anrede, Namen, technische Begriffe, Signatur und geschützte Platzhalter unverändert bei.

        Keine Erklärung, Überschrift, Einleitung, Abschlussbemerkung oder Markdown. Gib ausschließlich den überarbeiteten Antworttext zurück.
        """;
    private static readonly Regex ProtectedValuePattern = new(
        $@"{Regex.Escape(TicketCloseReplyComposer.CloseNotice)}|\$Customer(?:FirstName|LastName)?|\$Ticket(?:Number|Title)|\$Date|https?://[^\s<>]+|[\w.+-]+@[\w.-]+\.[A-Za-z]{{2,}}|(?:[A-Za-z]:\\|/)[^\s<>]+|\b(?:[A-Za-z]+[-_.])+(?:[A-Za-z0-9]+[-_.]?)*\d+[A-Za-z0-9_.-]*\b|\b0x[0-9A-Fa-f]{{6,}}\b",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex MetaLinePattern = new(
        @"(?im)^[ \t]*(?:ENTWURF|ENDE DES ENTWURFS|ÜBERARBEITETE VERSION):?[ \t]*(?:\r?\n|$)",
        RegexOptions.Compiled);
    private static readonly Regex IntroLinePattern = new(
        @"(?i)\A[ \t]*Hier ist die überarbeitete Version:[ \t]*(?:\r?\n)",
        RegexOptions.Compiled);

    private readonly IAiChatService _ai;
    private readonly LoggerService _logger;
    public bool CanRewrite => _ai.IsEnabled && _ai.CanChat;
    public string AvailabilityMessage => _ai.AvailabilityMessage;
    public event EventHandler? AvailabilityChanged;

    public AiTextRewriteService(IAiChatService ai, LoggerService logger)
    {
        _ai = ai; _logger = logger;
        _ai.StateChanged += (_, _) => AvailabilityChanged?.Invoke(this, EventArgs.Empty);
    }

    public async Task<string> RewriteCustomerReplyAsync(string draft, CancellationToken token = default)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(draft);
        if (draft.Length > MaximumTextLength) throw new ArgumentException("Der Antwortentwurf ist zu lang.", nameof(draft));
        if (!CanRewrite) throw new InvalidOperationException("Die KI ist derzeit nicht verfügbar.");

        var protectedValues = new List<string>();
        var masked = ProtectedValuePattern.Replace(draft, match =>
        {
            var marker = $"__PLENARO_PROTECTED_{protectedValues.Count}__";
            protectedValues.Add(match.Value); return marker;
        });
        var watch = Stopwatch.StartNew();
        _logger.OperationalInfo($"[AI Rewrite] action=start characters={draft.Length} provider={_ai.ProviderDescription}");
        try
        {
            var response = await _ai.ChatAsync(
                [new(AiChatRole.System, SystemPrompt), new(AiChatRole.User, masked)],
                new AiRequestOptions(0, CalculateMaxOutputTokens(draft)), token);
            if (string.IsNullOrWhiteSpace(response)) throw new InvalidDataException("Die KI hat keinen Text zurückgegeben.");
            var restored = CleanRewriteOutput(response);
            if (string.IsNullOrWhiteSpace(restored)) throw new InvalidDataException("Die KI hat keinen Text zurückgegeben.");
            for (var index = 0; index < protectedValues.Count; index++)
            {
                var marker = $"__PLENARO_PROTECTED_{index}__";
                if (!restored.Contains(marker, StringComparison.Ordinal))
                    throw new InvalidDataException("Ein geschützter Wert fehlt in der KI-Antwort.");
                restored = restored.Replace(marker, protectedValues[index], StringComparison.Ordinal);
            }
            restored = LimitLength(restored, MaximumTextLength);
            _logger.OperationalInfo($"[AI Rewrite] action=completed charactersIn={draft.Length} charactersOut={restored.Length} durationMs={watch.ElapsedMilliseconds}");
            return restored;
        }
        catch (Exception error) when (error is not OperationCanceledException)
        {
            _logger.Warning($"[AI Rewrite] action=failed charactersIn={draft.Length} durationMs={watch.ElapsedMilliseconds} errorType={error.GetType().Name}");
            throw;
        }
    }

    internal static int CalculateMaxOutputTokens(string text)
        => Math.Clamp(text.Length / 3 + 96, 128, 3072);

    internal static string CleanRewriteOutput(string response)
        => MetaLinePattern.Replace(IntroLinePattern.Replace(response, string.Empty), string.Empty).Trim();

    private static string LimitLength(string value, int maximum)
    {
        if (value.Length <= maximum) return value;
        var boundary = value.LastIndexOfAny([' ', '\r', '\n', '\t'], maximum - 1);
        return value[..(boundary > maximum / 2 ? boundary : maximum)].TrimEnd();
    }
}
