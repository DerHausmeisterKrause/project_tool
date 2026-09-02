using Microsoft.Data.Sqlite;
using TaskTool.Models;
using TaskTool.Services;
using Xunit;

namespace TaskTool.Tests;

public sealed class TicketDetailCacheServiceTests : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), $"plenaro-cache-{Guid.NewGuid():N}.db");
    private readonly TicketDetailCacheService _cache;

    public TicketDetailCacheServiceTests()
    {
        var database = new DatabaseService(new LoggerService(), _path);
        database.Initialize();
        _cache = new TicketDetailCacheService(database);
    }

    [Fact]
    public void NewAndMigratedEntriesAreIncompleteUntilFetchProfileProvesCompleteness()
    {
        _cache.Store(Context("1", Array.Empty<TicketArticleItem>(), "", ""), "open", DateTime.UtcNow);
        var entry = Assert.IsType<TicketDetailCacheEntry>(_cache.LoadEntry("1"));
        Assert.False(entry.ArticlesComplete);
        Assert.False(entry.DynamicFieldsComplete);
        Assert.False(entry.IsCompleteFor(20));
    }

    [Fact]
    public void IncreasingArticleLimitInvalidatesOnlyArticleCompletenessForRequestedView()
    {
        _cache.Store(Context("1", [Article("a")], "cost", "order"), "open", DateTime.UtcNow,
            TicketDetailFetchProfile.Full(20));
        var entry = _cache.LoadEntry("1")!;
        Assert.True(entry.MetadataComplete);
        Assert.True(entry.DynamicFieldsComplete);
        Assert.True(entry.IsCompleteFor(20));
        Assert.False(entry.IsCompleteFor(21));
    }

    [Fact]
    public void CandidateFetchCannotOverwriteRicherAssignedCache()
    {
        _cache.Store(Context("1", [Article("full")], "cost", "order"), "open", DateTime.UtcNow,
            TicketDetailFetchProfile.Full(20));
        _cache.Store(Context("1", [Article("candidate")], "", ""), "open", DateTime.UtcNow.AddMinutes(1),
            TicketDetailFetchProfile.Candidate(20));

        var entry = _cache.LoadEntry("1")!;
        Assert.True(entry.DynamicFieldsComplete);
        Assert.Equal("cost", entry.Context.CostCenterValue);
        Assert.Equal("full", Assert.Single(entry.Context.Articles).ArticleId);
    }

    [Fact]
    public void CandidateOnlyCacheRequiresLaterAssignedFetch()
    {
        _cache.Store(Context("1", [Article("candidate")], "", ""), "open", DateTime.UtcNow,
            TicketDetailFetchProfile.Candidate(20));
        var entry = _cache.LoadEntry("1")!;
        Assert.True(entry.ArticlesComplete);
        Assert.False(entry.DynamicFieldsComplete);
        Assert.False(entry.IsCompleteFor(20));
    }

    [Fact]
    public void EmptyOptionsDoNotDestroyLastValidFingerprintCache()
    {
        _cache.ReplaceFieldOptions("Cost", "fingerprint", [new TicketFieldOption("1", "One")]);
        _cache.ReplaceFieldOptions("Cost", "fingerprint", Array.Empty<TicketFieldOption>());
        Assert.Single(_cache.LoadFieldOptions("Cost", "fingerprint", TimeSpan.FromDays(1)));
        Assert.Empty(_cache.LoadFieldOptions("Cost", "different", TimeSpan.FromDays(1)));
    }

    [Fact]
    public void SchemaV22IsAdditiveAndMarksLegacyRowsIncomplete()
    {
        using var connection = new SqliteConnection($"Data Source={_path}"); connection.Open();
        using var version = connection.CreateCommand(); version.CommandText = "SELECT version FROM schema_version";
        Assert.Equal(22L, version.ExecuteScalar());
        using var columns = connection.CreateCommand(); columns.CommandText = "PRAGMA table_info(znuny_ticket_detail_cache)";
        using var reader = columns.ExecuteReader(); var names = new List<string>(); while (reader.Read()) names.Add(reader.GetString(1));
        Assert.Contains("articles_complete", names); Assert.Contains("dynamic_fields_complete", names);
    }

    private static TicketBookingContext Context(string id, IReadOnlyList<TicketArticleItem> articles, string cost, string order)
        => new(id, $"N{id}", cost, order, [], [], "", articles, articles.FirstOrDefault(), "customer@example.test", $"Ticket {id}");
    private static TicketArticleItem Article(string id) => new() { ArticleId = id, Subject = id, Body = id };
    public void Dispose() { SqliteConnection.ClearAllPools(); if (File.Exists(_path)) File.Delete(_path); }
}
