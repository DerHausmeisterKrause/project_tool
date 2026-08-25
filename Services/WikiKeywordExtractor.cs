using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using TaskTool.Models;

namespace TaskTool.Services;

public sealed class WikiKeywordExtractor
{
    private readonly WikiVocabularyIndexService? _vocabulary; private readonly LoggerService? _logger;
    private static readonly HashSet<string> StopWords = new(StringComparer.OrdinalIgnoreCase)
    { "hallo","zusammen","guten","morgen","abend","bitte","danke","vielen","freundlichen","grüßen","grüße","rückmeldung","ich","wir","uns","der","die","das","den","dem","des","ein","eine","einer","einem","einen","und","oder","aber","mit","auf","für","von","zu","ist","sind","war","werden","wurde","kann","können","nicht","nach","bei","über","wie","am","im","in","aus","an","the","and","or","please","thanks","this","that","with","from","hello","can","could","would","should","is","are","was","were","to","of","for","on","in","ticket","problem","prüfen","funktioniert" };
    private static readonly HashSet<string> Generic = new(StringComparer.OrdinalIgnoreCase) { "anleitung","dokumentation","information","system","benutzer","anwendung","problem","hilfe" };
    public WikiKeywordExtractor(WikiVocabularyIndexService? vocabulary = null, LoggerService? logger = null) { _vocabulary = vocabulary; _logger = logger; }

    public IReadOnlyList<string> Extract(string title, string firstMessage) => ExtractWeighted(title, firstMessage, Array.Empty<WikiVocabularyPage>()).Select(x => x.Text).ToList();
    public IReadOnlyList<WikiSearchTerm> ExtractForSource(WikiSourceSettings source, string title, string firstMessage)
    {
        var watch = Stopwatch.StartNew(); var seeds = Generate(title, firstMessage).Select(x => x.Text); var pages = _vocabulary?.FindCandidates(source, seeds) ?? Array.Empty<WikiVocabularyPage>(); var result = ExtractWeighted(title, firstMessage, pages);
        _logger?.Info($"[WikiExtract] sourceId={source.Id} candidateCount={Generate(title, firstMessage).Count} localIndexHits={pages.Count} selectedTermCount={result.Count} durationMs={watch.ElapsedMilliseconds}"); return result;
    }

    private static IReadOnlyList<WikiSearchTerm> ExtractWeighted(string title, string message, IReadOnlyList<WikiVocabularyPage> pages)
    {
        var candidates = Generate(title, message); var ticketNormalized = Normalize(title + " " + CleanMessage(message)); var ticketTokens = Tokenize(ticketNormalized).ToHashSet();
        foreach (var page in pages)
        {
            var normalized = Normalize(page.Title); var similarity = Dice(ticketTokens, Tokenize(normalized).ToHashSet()); if (similarity < .25) continue;
            foreach (var phrase in PhraseParts(page.Title)) Add(candidates, phrase, 35 + similarity * 45, "WikiVocabulary", true, similarity);
        }
        var pageCount = Math.Max(1, pages.Count);
        foreach (var c in candidates.Values) { var df = pages.Count(p => Normalize(p.Title).Contains(c.NormalizedText, StringComparison.Ordinal)); c.Idf = Math.Clamp(Math.Log((pageCount + 1d) / (df + 1d)) * 10, 0, 30); c.Score += c.Idf - (Generic.Contains(c.Text) ? 25 : 0); }
        var selected = new List<Candidate>();
        foreach (var candidate in candidates.Values.Where(x => x.Score >= 8).OrderByDescending(x => x.Score)) { if (selected.Any(x => Redundant(x.NormalizedText, candidate.NormalizedText))) continue; selected.Add(candidate); if (selected.Count == 6) break; }
        return selected.Select(x => new WikiSearchTerm(x.Text, x.NormalizedText, Math.Clamp(x.Score, 0, 100), x.Origin, x.Text.Contains(' '), x.Origin == "WikiVocabulary", x.Similarity, x.Idf)).ToList();
    }

    private static Dictionary<string, Candidate> Generate(string title, string message)
    {
        var result = new Dictionary<string, Candidate>(StringComparer.Ordinal); AddNgrams(result, title, 4, 35, "Title"); AddNgrams(result, CleanMessage(message), 3, 7, "Message");
        foreach (var phrase in PhraseParts(title + ". " + CleanMessage(message))) Add(result, phrase, 15 + Tokenize(phrase).Count() * 3, "Rake", true);
        foreach (var raw in TokenizeOriginal(title + " " + message).Where(IsStructured)) Add(result, raw, 20, "Structured", false);
        return result;
    }
    private static void AddNgrams(Dictionary<string, Candidate> result, string text, int maxN, double baseScore, string origin) { var words = TokenizeOriginal(text).Where(x => !StopWords.Contains(Normalize(x))).ToArray(); for (var n = 1; n <= maxN; n++) for (var i = 0; i + n <= words.Length; i++) Add(result, string.Join(" ", words.Skip(i).Take(n)), baseScore + n * 5, origin, n > 1); }
    private static IEnumerable<string> PhraseParts(string text) => Regex.Split(text, @"[\r\n,;.!?]|\b(?:und|oder|and|or|aber|but|bitte|please)\b", RegexOptions.IgnoreCase).Select(x => string.Join(" ", TokenizeOriginal(x).Where(w => !StopWords.Contains(Normalize(w))).Take(4))).Where(x => Tokenize(x).Any());
    private static void Add(Dictionary<string, Candidate> result, string text, double score, string origin, bool phrase, double similarity = 0) { text = Regex.Replace(text.Trim(), @"\s+", " "); var normalized = Normalize(text); if (normalized.Length < 3 || StopWords.Contains(normalized)) return; if (!result.TryGetValue(normalized, out var c)) result[normalized] = new(text, normalized, score, origin, similarity); else { c.Score = Math.Max(c.Score, score) + 5; c.Similarity = Math.Max(c.Similarity, similarity); } }
    public static IEnumerable<string> Tokenize(string text) => Regex.Matches(Normalize(text), @"[\p{L}\p{N}][\p{L}\p{N}._-]*").Select(x => x.Value);
    private static IEnumerable<string> TokenizeOriginal(string text) => Regex.Matches(text ?? "", @"[\p{L}][\p{L}\p{N}._-]*|\d{1,5}").Select(x => x.Value);
    public static string Normalize(string text) { text = (text ?? "").Normalize(NormalizationForm.FormD).ToLowerInvariant().Replace("ß", "ss"); var b = new StringBuilder(); foreach (var c in text) if (CharUnicodeInfo.GetUnicodeCategory(c) != UnicodeCategory.NonSpacingMark) b.Append(char.IsWhiteSpace(c) || c is '_' or '-' ? ' ' : c); return Regex.Replace(b.ToString().Normalize(NormalizationForm.FormC), @"\s+", " ").Trim(); }
    private static string CleanMessage(string text) { text = Regex.Replace(text ?? "", @"(?im)^(from|von|to|an|subject|betreff|sent|gesendet):.*$|^>.*$", " "); text = Regex.Replace(text, @"(?is)(mit freundlichen grüßen|viele grüße|best regards).*$", " "); return Regex.Replace(text, @"https?://\S+|\b[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}\b|\b(?:ticket|tn)?#?\d{5,}\b", " ", RegexOptions.IgnoreCase); }
    private static bool IsStructured(string x) => Regex.IsMatch(x, @"\d|[-_.]|[A-Z]{2,}|[a-z][A-Z]");
    private static double Dice(HashSet<string> a, HashSet<string> b) => a.Count + b.Count == 0 ? 0 : 2d * a.Intersect(b).Count() / (a.Count + b.Count);
    private static bool Redundant(string a, string b) { if (a == b) return true; var aa = Tokenize(a).ToHashSet(); var bb = Tokenize(b).ToHashSet(); var intersection = aa.Intersect(bb).Count(); if (Math.Min(aa.Count, bb.Count) == 1 && intersection == 1) return true; return aa.Count > 0 && (double)intersection / aa.Union(bb).Count() >= .6; }
    private sealed class Candidate(string text, string normalized, double score, string origin, double similarity) { public string Text=text; public string NormalizedText=normalized; public double Score=score; public string Origin=origin; public double Similarity=similarity; public double Idf; }
}
