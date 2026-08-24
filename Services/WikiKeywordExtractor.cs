using System.Text.RegularExpressions;

namespace TaskTool.Services;

public sealed class WikiKeywordExtractor
{
    private static readonly HashSet<string> StopWords = new(StringComparer.OrdinalIgnoreCase)
    { "hallo", "bitte", "danke", "ich", "wir", "der", "die", "das", "den", "dem", "ein", "eine", "und", "oder", "mit", "auf", "für", "von", "zu", "ist", "sind", "the", "and", "or", "please", "thanks", "this", "that", "with", "from", "hello", "aktualisieren", "problem", "ticket" };

    public IReadOnlyList<string> Extract(string title, string firstMessage)
    {
        var candidates = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        Add(title, 5, candidates);
        Add(firstMessage, 1, candidates);
        return candidates.Where(x => x.Key.Length >= 3)
            .OrderByDescending(x => x.Value).ThenByDescending(x => x.Key.Contains(' ')).ThenBy(x => x.Key)
            .Select(x => x.Key).Distinct(StringComparer.OrdinalIgnoreCase).Take(8).ToList();
    }

    private static void Add(string? text, int weight, Dictionary<string, int> candidates)
    {
        if (string.IsNullOrWhiteSpace(text)) return;
        text = Regex.Replace(text, @"\b[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}\b|\b(?:ticket|tn)?#?\d{5,}\b", " ", RegexOptions.IgnoreCase);
        var words = Regex.Matches(text, @"[\p{L}][\p{L}\p{N}._-]*|\d{1,4}").Select(m => m.Value).ToArray();
        for (var i = 0; i < words.Length; i++)
        {
            var word = words[i].Trim('.', '_', '-');
            if (word.Length < 3 || StopWords.Contains(word)) continue;
            var term = word;
            if (i + 1 < words.Length && Regex.IsMatch(words[i + 1], @"^\d{1,4}$") && !Regex.IsMatch(word, @"^\d"))
                term += " " + words[i + 1];
            else if (i + 1 < words.Length && char.IsUpper(word[0]) && char.IsUpper(words[i + 1][0]) && !StopWords.Contains(words[i + 1]))
                term += " " + words[i + 1];
            candidates[term] = candidates.GetValueOrDefault(term) + weight + (term.Contains(' ') ? 2 : 0) + (Regex.IsMatch(term, @"[-.]|\d") ? 2 : 0);
        }
    }
}
