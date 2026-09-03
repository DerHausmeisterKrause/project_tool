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
        Assert.True(ZnunySyncPolicy.RequiresFullTicketGet(entry, 20));
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
    public void SameVersionCandidateFetchPreservesCompleteDynamicFields()
    {
        var changed = new DateTime(2026, 1, 1, 10, 0, 0, DateTimeKind.Utc);
        _cache.Store(Context("1", [Article("full")], "cost", "order"), "open", changed,
            TicketDetailFetchProfile.Full(20));
        _cache.Store(Context("1", [Article("candidate")], "", ""), "open", changed,
            TicketDetailFetchProfile.Candidate(20));

        var entry = _cache.LoadEntry("1")!;
        Assert.True(entry.DynamicFieldsComplete);
        Assert.Equal("cost", entry.Context.CostCenterValue);
        Assert.Equal("candidate", Assert.Single(entry.Context.Articles).ArticleId);
        Assert.True(entry.IsCompleteFor(20));
    }

    [Fact]
    public void NewVersionCandidateFetchInvalidatesUnfetchedDynamicFields()
    {
        var oldChanged = new DateTime(2026, 1, 1, 10, 0, 0, DateTimeKind.Utc);
        var newChanged = oldChanged.AddHours(1);
        _cache.Store(Context("1", [Article("old")], "cost", "order"), "open", oldChanged,
            TicketDetailFetchProfile.Full(20));
        _cache.Store(Context("1", [Article("new")], "", ""), "open", newChanged,
            TicketDetailFetchProfile.Candidate(20));

        var entry = _cache.LoadEntry("1")!;
        Assert.Equal("cost", entry.Context.CostCenterValue); // last-known-good display fallback
        Assert.False(entry.DynamicFieldsComplete);
        Assert.True(entry.ArticlesComplete);
        Assert.Equal(20, entry.FetchedArticleLimit);
        Assert.Equal("new", Assert.Single(entry.Context.Articles).ArticleId);
        Assert.False(entry.IsCompleteFor(20));
        Assert.True(ZnunySyncPolicy.RequiresFullTicketGet(entry, 20));

        _cache.Store(Context("1", [Article("assigned")], "new-cost", "new-order"), "open", newChanged,
            TicketDetailFetchProfile.Full(20));
        var repaired = _cache.LoadEntry("1")!;
        Assert.False(ZnunySyncPolicy.RequiresFullTicketGet(repaired, 20));
        Assert.True(repaired.IsCompleteFor(20));
    }

    [Fact]
    public void NewVersionWithoutArticlesInvalidatesArticleVersionAndLimit()
    {
        var oldChanged = new DateTime(2026, 1, 1, 10, 0, 0, DateTimeKind.Utc);
        _cache.Store(Context("1", [Article("old")], "cost", "order"), "open", oldChanged,
            TicketDetailFetchProfile.Full(20));
        _cache.Store(Context("1", [], "new-cost", "new-order"), "open", oldChanged.AddHours(1),
            new TicketDetailFetchProfile(true, false, true, 0));

        var entry = _cache.LoadEntry("1")!;
        Assert.Equal("old", Assert.Single(entry.Context.Articles).ArticleId);
        Assert.False(entry.ArticlesComplete);
        Assert.Equal(0, entry.FetchedArticleLimit);
        Assert.False(entry.IsCompleteFor(20));
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
    public void SchemaV25IsAdditiveAndKeepsPreviousMigrations()
    {
        using var connection = new SqliteConnection($"Data Source={_path}"); connection.Open();
        using var version = connection.CreateCommand(); version.CommandText = "SELECT version FROM schema_version";
        Assert.Equal(25L, version.ExecuteScalar());
        using var columns = connection.CreateCommand(); columns.CommandText = "PRAGMA table_info(znuny_ticket_detail_cache)";
        using var reader = columns.ExecuteReader(); var names = new List<string>(); while (reader.Read()) names.Add(reader.GetString(1));
        Assert.Contains("articles_complete", names); Assert.Contains("dynamic_fields_complete", names);
        reader.Close();
        using var poolTable = connection.CreateCommand();
        poolTable.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='znuny_candidate_pool_snapshot'";
        Assert.Equal(1L, poolTable.ExecuteScalar());
        using var readTables = connection.CreateCommand();
        readTables.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name IN ('znuny_ticket_article_read_baseline','znuny_ticket_article_read_state')";
        Assert.Equal(2L, readTables.ExecuteScalar());
    }

    [Theory]
    [InlineData(150)]
    [InlineData(300)]
    public void PersistentReconciliationCycleConvergesAcrossBudgetedPasses(int ticketCount)
    {
        var ids = Enumerable.Range(1, ticketCount).Select(value => value.ToString()).ToList();
        var processed = new HashSet<string>();
        var runs = 0;
        while (true)
        {
            var cycle = _cache.StartOrLoadCycle("agent", "stable-discovery", ids);
            if (cycle.PendingTicketIds.Count == 0) break;
            foreach (var id in cycle.PendingTicketIds.Take(50))
            {
                Assert.True(processed.Add(id));
                _cache.CompleteCycleTicket("agent", id);
            }
            runs++;
            Assert.True(runs <= Math.Ceiling(ticketCount / 50d));
        }

        Assert.Equal(ticketCount, processed.Count);
        Assert.True(runs > 1);
        _cache.CompleteCycle("agent");
        Assert.Equal(ticketCount, _cache.StartOrLoadCycle("agent", "stable-discovery", ids).PendingTicketIds.Count);
    }

    [Fact]
    public void ChangedDiscoverySafelyStartsANewCycle()
    {
        var first = _cache.StartOrLoadCycle("agent", "one", ["1", "2"]);
        _cache.CompleteCycleTicket("agent", "1");
        var changed = _cache.StartOrLoadCycle("agent", "two", ["2", "3"]);
        Assert.Equal(["2", "3"], changed.PendingTicketIds);
    }

    [Fact]
    public void ChangedDiscoveryRetainsCompletedTicketsThatAreStillPresent()
    {
        _cache.StartOrLoadCycle("agent", "one", ["1", "2"]);
        _cache.CompleteCycleTicket("agent", "1");

        var changed = _cache.StartOrLoadCycle("agent", "two", ["1", "2", "3"]);

        Assert.Equal(["2", "3"], changed.PendingTicketIds);
        Assert.Equal(["1", "2", "3"], changed.DiscoveredTicketIds);
    }

    [Theory]
    [InlineData(ZnunyReconciliationWorkKind.RemovalVerification, ZnunyReconciliationWorkKind.Assigned, true)]
    [InlineData(ZnunyReconciliationWorkKind.Assigned, ZnunyReconciliationWorkKind.RemovalVerification, true)]
    [InlineData(ZnunyReconciliationWorkKind.Assigned, ZnunyReconciliationWorkKind.Assigned, false)]
    [InlineData(ZnunyReconciliationWorkKind.RemovalVerification, ZnunyReconciliationWorkKind.RemovalVerification, false)]
    public void ChangedDiscoveryRetainsCompletionOnlyForSameWorkKind(
        ZnunyReconciliationWorkKind firstKind, ZnunyReconciliationWorkKind secondKind, bool expectedPending)
    {
        var first = new ZnunyReconciliationWorkItem("42", firstKind).PersistedKey;
        var second = new ZnunyReconciliationWorkItem("42", secondKind).PersistedKey;
        _cache.StartOrLoadCycle("agent", "one", [first]);
        _cache.CompleteCycleTicket("agent", first);

        var changed = _cache.StartOrLoadCycle("agent", "two", [second]);

        Assert.Equal(expectedPending ? [second] : [], changed.PendingTicketIds);
    }

    [Fact]
    public void RepeatedDiscoveryChangesDoNotStarvePendingWork()
    {
        var completed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var generation = 0; generation < 20; generation++)
        {
            var items = Enumerable.Range(generation, 150)
                .Select(id => new ZnunyReconciliationWorkItem(id.ToString(), ZnunyReconciliationWorkKind.Assigned).PersistedKey)
                .ToList();
            var cycle = _cache.StartOrLoadCycle("agent", $"generation-{generation}", items);
            foreach (var item in cycle.PendingTicketIds.Take(50))
            {
                completed.Add(item);
                _cache.CompleteCycleTicket("agent", item);
            }
        }

        var finalItems = Enumerable.Range(19, 150)
            .Select(id => new ZnunyReconciliationWorkItem(id.ToString(), ZnunyReconciliationWorkKind.Assigned).PersistedKey)
            .ToList();
        var finalCycle = _cache.StartOrLoadCycle("agent", "generation-19", finalItems);
        while (finalCycle.PendingTicketIds.Count > 0)
        {
            foreach (var item in finalCycle.PendingTicketIds.Take(50))
            {
                completed.Add(item);
                _cache.CompleteCycleTicket("agent", item);
            }
            finalCycle = _cache.StartOrLoadCycle("agent", "generation-19", finalItems);
        }

        Assert.Empty(finalCycle.PendingTicketIds);
        Assert.All(finalItems, item => Assert.Contains(item, completed));
    }

    [Fact]
    public void PersistedDynamicFieldFreshnessUsesOriginalFetchedTimestamp()
    {
        _cache.ReplaceFieldOptions("Cost", "fingerprint", [new TicketFieldOption("1", "One")]);
        using (var connection = new SqliteConnection($"Data Source={_path}"))
        {
            connection.Open(); using var age = connection.CreateCommand();
            age.CommandText = "UPDATE znuny_dynamic_field_options_cache SET fetched_utc=$utc";
            age.Parameters.AddWithValue("$utc", DateTime.UtcNow.AddHours(-23).ToString("O")); age.ExecuteNonQuery();
        }
        var entry = _cache.LoadFieldOptionsEntry("Cost", "fingerprint");
        Assert.True(entry.IsFresh(TimeSpan.FromHours(24), DateTime.UtcNow));
        Assert.False(entry.IsFresh(TimeSpan.FromHours(22), DateTime.UtcNow));
    }

    private static TicketBookingContext Context(string id, IReadOnlyList<TicketArticleItem> articles, string cost, string order)
        => new(id, $"N{id}", cost, order, [], [], "", articles, articles.FirstOrDefault(), "customer@example.test", $"Ticket {id}");
    private static TicketArticleItem Article(string id) => new() { ArticleId = id, Subject = id, Body = id };
    public void Dispose() { SqliteConnection.ClearAllPools(); if (File.Exists(_path)) File.Delete(_path); }
}
