using TaskTool.Models;
using Xunit;

namespace TaskTool.Tests;

public sealed class DurationTextParserTests
{
    [Theory]
    [InlineData("03:30:00", 12600)]
    [InlineData("27:15:00", 98100)]
    [InlineData("100:00:00", 360000)]
    public void ParsesUnboundedHours(string text, long expected)
    {
        Assert.True(DurationTextParser.TryParseSeconds(text, out var actual));
        Assert.Equal(expected, actual);
    }

    [Theory]
    [InlineData("03:60:00")]
    [InlineData("03:00:60")]
    [InlineData("-1:00:00")]
    [InlineData("abc")]
    public void RejectsInvalidDurations(string text)
        => Assert.False(DurationTextParser.TryParseSeconds(text, out _));
}
