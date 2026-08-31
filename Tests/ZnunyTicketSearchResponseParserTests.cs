using TaskTool.Models;
using Xunit;

namespace TaskTool.Tests;

public sealed class ZnunyTicketSearchResponseParserTests
{
    [Fact]
    public void EmptyObjectIsSuccessfulZeroResult()
    {
        var result = Parse("{}");

        Assert.Empty(result.TicketIds);
        Assert.Equal(ZnunyTicketSearchResponseShape.EmptyObject, result.ResponseShape);
    }

    [Theory]
    [InlineData("{\"TicketID\":[]}")]
    [InlineData("{\"TicketIDs\":[]}")]
    public void EmptyTicketIdArrayIsSuccessfulZeroResult(string json)
        => Assert.Empty(Parse(json).TicketIds);

    [Theory]
    [InlineData("{\"TicketID\":[\"123\",\"456\"]}")]
    [InlineData("{\"TicketIDs\":[\"123\",\"456\"]}")]
    public void TicketIdArraysAreExtracted(string json)
        => Assert.Equal(new[] { "123", "456" }, Parse(json).TicketIds);

    [Fact]
    public void ScalarTicketIdIsExtracted()
        => Assert.Equal(new[] { "123" }, Parse("{\"TicketID\":\"123\"}").TicketIds);

    [Fact]
    public void ZnunyErrorObjectIsNotTreatedAsZeroResults()
    {
        var exception = Assert.Throws<ZnunyApiException>(() =>
            Parse("{\"Error\":{\"ErrorCode\":\"TicketSearch.AuthFail\",\"ErrorMessage\":\"Authentication failed\"}}"));

        Assert.Equal("TicketSearch.AuthFail", exception.ErrorCode);
        Assert.NotEqual("Protocol", exception.ErrorCode);
    }

    [Theory]
    [InlineData("{\"foo\":\"bar\"}")]
    [InlineData("invalid JSON")]
    [InlineData("[]")]
    [InlineData("\"foo\"")]
    [InlineData("123")]
    public void UnknownOrInvalidResponseShapeRemainsProtocolError(string json)
    {
        var exception = Assert.Throws<ZnunyApiException>(() => Parse(json));

        Assert.Equal("Protocol", exception.ErrorCode);
    }

    [Fact]
    public void OpenAndNewPartialSearchesMergeWithoutDuplicates()
    {
        var owner = ZnunySyncPolicy.MergeTicketIds(
            Parse("{\"TicketID\":[\"1\",\"2\"]}").TicketIds,
            Parse("{}").TicketIds);
        var responsible = ZnunySyncPolicy.MergeTicketIds(
            Parse("{\"TicketID\":[\"2\",\"3\"]}").TicketIds,
            Parse("{}").TicketIds);
        var unique = ZnunySyncPolicy.MergeTicketIds(owner, responsible);

        Assert.Equal(new[] { "1", "2" }, owner);
        Assert.Equal(new[] { "2", "3" }, responsible);
        Assert.Equal(new[] { "1", "2", "3" }, unique);
    }

    [Fact]
    public void ConnectionSearchesCanBothReturnEmptyObjects()
    {
        var owner = Parse("{}");
        var responsible = Parse("{}");

        Assert.Empty(owner.TicketIds);
        Assert.Empty(responsible.TicketIds);
    }

    private static ZnunyTicketSearchParseResult Parse(string json)
        => ZnunyTicketSearchResponseParser.ExtractTicketIdsStrict(json, "TicketSearchTest");
}
