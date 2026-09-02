using Microsoft.Data.Sqlite;
using TaskTool.Models;
using TaskTool.Services;
using Xunit;

namespace TaskTool.Tests;

public sealed class TicketCandidateSnapshotServiceTests : IDisposable
{
    private readonly string _path = Path.Combine(Path.GetTempPath(), $"plenaro-candidate-pool-{Guid.NewGuid():N}.db");
    private readonly TicketCandidateSnapshotService _snapshots;

    public TicketCandidateSnapshotServiceTests()
    {
        var database = new DatabaseService(new LoggerService(), _path);
        database.Initialize();
        _snapshots = new TicketCandidateSnapshotService(database);
    }

    [Fact]
    public void FilteredAndPoolSnapshotsKeepSeparateSemanticsAcrossRestart()
    {
        var matched = Ticket("1", "keyword");
        var unmatched = Ticket("2", "");
        _snapshots.Replace([matched]);
        _snapshots.ReplacePool([matched, unmatched]);

        var restarted = new TicketCandidateSnapshotService(Database());
        Assert.Equal(new[] { "1" }, restarted.Load().Select(item => item.TicketId));
        Assert.Equal(new[] { "2", "1" }, restarted.LoadPool().Select(item => item.TicketId));
    }

    [Fact]
    public void ReplacingFilteredSnapshotCannotEraseStalePoolFallback()
    {
        _snapshots.ReplacePool([Ticket("9", "")]);
        _snapshots.Replace([]);
        Assert.Equal("9", _snapshots.LoadPool().Single().TicketId);
    }

    [Fact]
    public void PoolReplacementPersistsSuccessfulRefreshAndLocalRemoval()
    {
        _snapshots.ReplacePool([Ticket("1", "keyword"), Ticket("2", "")]);
        _snapshots.ReplacePool(_snapshots.LoadPool().Where(item => item.TicketId != "2").ToList());
        Assert.Equal("1", new TicketCandidateSnapshotService(Database()).LoadPool().Single().TicketId);
    }

    private DatabaseService Database()
    {
        var database = new DatabaseService(new LoggerService(), _path);
        database.Initialize();
        return database;
    }

    private static ZnunyCandidateTicket Ticket(string id, string keyword) => new()
    {
        TicketId = id, TicketNumber = id, Title = "Ticket " + id, DescriptionPreview = "Preview",
        Owner = "Pool", Responsible = "Pool", State = "open", WebUrl = "https://znuny/" + id,
        MatchedKeyword = keyword, LastSyncedUtc = DateTime.UtcNow
    };

    public void Dispose()
    {
        SqliteConnection.ClearAllPools();
        if (File.Exists(_path)) File.Delete(_path);
        Assert.False(File.Exists(_path));
    }
}
