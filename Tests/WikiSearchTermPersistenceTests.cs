using TaskTool.Models;
using Xunit;

namespace TaskTool.Tests;

public sealed class WikiSearchTermPersistenceTests
{
    [Fact]
    public void StoredTermsAreRestoredInTheirOriginalOrder()
    {
        var terms = WikiSearchTermPersistence.MergeSerialized(
            new[] { "[\"Debian\",\"Confluence\"]" });

        Assert.Equal(new[] { "Debian", "Confluence" }, terms);
    }

    [Fact]
    public void TermsFromMultipleSuccessfulRunsAreMergedCaseInsensitively()
    {
        var terms = WikiSearchTermPersistence.MergeSerialized(new[]
        {
            "[\"Debian\",\"Confluence\"]",
            "[\"debian\",\"PostgreSQL\"]"
        });

        Assert.Equal(new[] { "Debian", "Confluence", "PostgreSQL" }, terms);
    }

    [Fact]
    public void InvalidAndEmptyJsonCannotBreakTaskDetails()
    {
        var terms = WikiSearchTermPersistence.MergeSerialized(new string?[]
        {
            "not-json", null, string.Empty, "{}", "[\"PostgreSQL\"]"
        });

        Assert.Equal(new[] { "PostgreSQL" }, terms);
    }

    [Fact]
    public void DisplayTermsAreLimitedToSix()
    {
        var terms = WikiSearchTermPersistence.MergeSerialized(
            new[] { "[\"one\",\"two\",\"three\",\"four\",\"five\",\"six\",\"seven\"]" });

        Assert.Equal(6, terms.Count);
        Assert.DoesNotContain("seven", terms);
    }
}
