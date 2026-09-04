using Microsoft.Data.Sqlite;
using TaskTool.Models;
using TaskTool.Services;
using Xunit;

namespace TaskTool.Tests;

public sealed class TicketTimeBookingAccountingTests : IDisposable
{
    private readonly string _databasePath = Path.Combine(Path.GetTempPath(), $"plenaro-booking-{Guid.NewGuid():N}.db");
    private readonly string _settingsPath = Path.Combine(Path.GetTempPath(), $"plenaro-booking-settings-{Guid.NewGuid():N}.json");
    private readonly DatabaseService _database;
    private readonly TaskService _tasks;
    private readonly Guid _taskId;

    public TicketTimeBookingAccountingTests()
    {
        var logger = new LoggerService();
        var settings = new SettingsService(logger, _settingsPath);
        settings.Current.OutlookSyncEnabled = false;
        _database = new DatabaseService(logger, _databasePath);
        _database.Initialize();
        _tasks = new TaskService(_database, logger, new OutlookInteropService(logger, settings), settings);
        _taskId = _tasks.CreateTask(new TaskItem { Title = "Ticket" }).Id;
    }

    [Theory]
    [InlineData(515, 15)]
    [InlineData(900, 15)]
    [InlineData(901, 30)]
    [InlineData(1800, 30)]
    [InlineData(4200, 75)]
    [InlineData(4500, 75)]
    [InlineData(4560, 90)]
    public void BookingMinutesAreRoundedUpToFullFifteenMinuteIntervals(long seconds, int expectedMinutes)
        => Assert.Equal(expectedMinutes, TicketSystemService.CalculateBookedMinutes(seconds));

    [Fact]
    public void CompletionNormalizesDisplayedMinutesButPreservesExactSourceSeconds()
    {
        var booking = Booking("Pending", 70, 75, 4200);
        _tasks.CreateTicketTimeBooking(booking);

        var pending = _tasks.GetPendingTicketTimeBooking(_taskId)!;
        Assert.Equal(70, pending.Minutes);
        Assert.Equal(75, pending.BookedMinutes);
        Assert.Equal(4200, pending.SourceSeconds);

        _tasks.CompleteTicketTimeBooking(pending, "article");

        Assert.Equal(75, pending.Minutes);
        Assert.Equal(75, pending.BookedMinutes);
        Assert.Equal(4200, pending.SourceSeconds);
        var persisted = _tasks.GetSuccessfulTicketTimeBookings(_taskId).Single();
        Assert.Equal(75, persisted.Minutes);
        Assert.Equal(75, persisted.BookedMinutes);
        Assert.Equal(4200, persisted.SourceSeconds);
        Assert.Equal(4200, _tasks.GetSuccessfullyTransferredSeconds(_taskId));
    }

    [Fact]
    public void SuccessfulStatisticsUseBookedMinutesAndExcludeOtherStatuses()
    {
        _tasks.CreateTicketTimeBooking(Booking("Succeeded", 8.58m, 15, 515));
        _tasks.CreateTicketTimeBooking(Booking("Succeeded", 70, 75, 4200));
        _tasks.CreateTicketTimeBooking(Booking("Pending", 70, 75, 4200));
        _tasks.CreateTicketTimeBooking(Booking("Failed", 70, 75, 4200));

        var statistics = _tasks.GetSuccessfulBookingStatistics(DateTime.UtcNow.Date, TimeZoneInfo.Utc);

        Assert.Equal(90 * 60, statistics.TodaySeconds);
        Assert.Equal(90 * 60, statistics.SecondsByMonth[new DateTime(DateTime.UtcNow.Year, DateTime.UtcNow.Month, 1)]);
        Assert.Equal(515 + 4200, _tasks.GetSuccessfullyTransferredSeconds(_taskId));
    }

    [Fact]
    public void SuccessfulStatisticsReportSeventyFiveBookedMinutesInsteadOfSeventySourceMinutes()
    {
        _tasks.CreateTicketTimeBooking(Booking("Succeeded", 70, 75, 4200));

        var statistics = _tasks.GetSuccessfulBookingStatistics(DateTime.UtcNow.Date, TimeZoneInfo.Utc);

        Assert.Equal(4500, statistics.TodaySeconds);
        Assert.Equal(4200, _tasks.GetSuccessfullyTransferredSeconds(_taskId));
    }

    [Fact]
    public void MigrationV26NormalizesOnlySucceededRowsAndKeepsV25Structures()
    {
        _tasks.CreateTicketTimeBooking(Booking("Succeeded", 70, 75, 4200));
        _tasks.CreateTicketTimeBooking(Booking("Pending", 70, 75, 4200));
        _tasks.CreateTicketTimeBooking(Booking("Failed", 70, 75, 4200));
        Execute("UPDATE schema_version SET version=25");

        _database.Initialize();
        _database.Initialize();

        using var connection = new SqliteConnection(_database.ConnectionString);
        connection.Open();
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT status, minutes, booked_minutes, source_seconds FROM ticket_time_bookings ORDER BY status";
        using var reader = command.ExecuteReader();
        var rows = new Dictionary<string, (decimal Minutes, decimal BookedMinutes, long SourceSeconds)>();
        while (reader.Read())
            rows.Add(reader.GetString(0), (reader.GetDecimal(1), reader.GetDecimal(2), reader.GetInt64(3)));

        Assert.Equal((70m, 75m, 4200L), rows["Failed"]);
        Assert.Equal((70m, 75m, 4200L), rows["Pending"]);
        Assert.Equal((75m, 75m, 4200L), rows["Succeeded"]);
        reader.Close();
        command.CommandText = "SELECT version FROM schema_version";
        Assert.Equal(26L, command.ExecuteScalar());
        command.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name IN ('znuny_ticket_article_read_baseline','znuny_ticket_article_read_state')";
        Assert.Equal(2L, command.ExecuteScalar());
    }

    private TicketTimeBooking Booking(string status, decimal minutes, decimal bookedMinutes, long sourceSeconds) => new()
    {
        TaskId = _taskId,
        TicketId = "42",
        TicketNumber = "20260042",
        Minutes = minutes,
        BookedMinutes = bookedMinutes,
        SourceSeconds = sourceSeconds,
        BookedAtUtc = DateTime.UtcNow,
        Status = status
    };

    private void Execute(string sql)
    {
        using var connection = new SqliteConnection(_database.ConnectionString);
        connection.Open();
        using var command = connection.CreateCommand();
        command.CommandText = sql;
        command.ExecuteNonQuery();
    }

    public void Dispose()
    {
        SqliteConnection.ClearAllPools();
        if (File.Exists(_databasePath)) File.Delete(_databasePath);
        if (File.Exists(_settingsPath)) File.Delete(_settingsPath);
    }
}
