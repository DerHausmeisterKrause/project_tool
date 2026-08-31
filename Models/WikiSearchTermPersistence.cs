using System.Text.Json;

namespace TaskTool.Models;

internal static class WikiSearchTermPersistence
{
    internal const int MaximumDisplayTerms = 6;

    internal static IReadOnlyList<string> MergeSerialized(IEnumerable<string?> serializedRuns)
    {
        var terms = new List<string>(MaximumDisplayTerms);
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var serialized in serializedRuns)
        {
            if (string.IsNullOrWhiteSpace(serialized))
                continue;

            try
            {
                var storedTerms = JsonSerializer.Deserialize<string[]>(serialized);
                if (storedTerms == null)
                    continue;

                foreach (var storedTerm in storedTerms)
                {
                    var term = storedTerm?.Trim();
                    if (string.IsNullOrWhiteSpace(term) || !seen.Add(term))
                        continue;

                    terms.Add(term);
                    if (terms.Count == MaximumDisplayTerms)
                        return terms;
                }
            }
            catch (JsonException)
            {
                // A damaged historic row must not prevent the task details from opening.
            }
        }

        return terms;
    }
}
