using Microsoft.Data.Sqlite;
using TaskTool.Models;
using TaskTool.Services;
using TaskStatus = TaskTool.Models.TaskStatus;
using Xunit;

namespace TaskTool.Tests;

public sealed class TaskServiceUnbookedTimeTests : IDisposable
{
    private readonly string _databasePath = Path.Combine(Path.GetTempPath(), $"plenaro-time-edit-{Guid.NewGuid():N}.db");
    private readonly string _settingsPath = Path.Combine(Path.GetTempPath(), $"plenaro-time-edit-settings-{Guid.NewGuid():N}.json");
    private readonly DatabaseService _database;
    private readonly TaskService _tasks;

    public TaskServiceUnbookedTimeTests()
    {
        var logger = new LoggerService();
        var settings = new SettingsService(logger, _settingsPath);
        settings.Current.OutlookSyncEnabled = false;
        _database = new DatabaseService(logger, _databasePath);
        _database.Initialize();
        _tasks = new TaskService(_database, logger, new OutlookInteropService(logger, settings), settings);
    }

    [Fact]
    public void SetUnbookedTimeChangesOnlyCumulativeLocalSeconds()
    {
        var task = _tasks.CreateTask(new TaskItem { Title = "Ticket", TicketSecondsBooked = 600 });
        var booking = Booking(task.Id);
        _tasks.CreateTicketTimeBooking(booking);

        _tasks.SetUnbookedTicketSeconds(task, 14400, 300, booking.SourceSeconds);

        Assert.Equal(300 + booking.SourceSeconds + 14400, task.TicketSecondsBooked);
        var persistedBooking = _tasks.GetAllTicketTimeBookings(task.Id).Single();
        Assert.Equal(booking.Id, persistedBooking.Id);
        Assert.Equal(booking.BookingId, persistedBooking.BookingId);
        Assert.Equal(booking.ArticleId, persistedBooking.ArticleId);
        Assert.Equal(booking.SourceSeconds, persistedBooking.SourceSeconds);
        Assert.Equal(booking.BookedMinutes, persistedBooking.BookedMinutes);
        Assert.Equal(booking.CostCenter, persistedBooking.CostCenter);
        Assert.Equal(booking.Order, persistedBooking.Order);
        Assert.Equal(0, _tasks.GetTicketTimeBookingBaselineSeconds(task.Id));
    }

    [Fact]
    public void RunningTimerIsStoppedAdjustedAndRestartedWithoutDuplicateOpenLog()
    {
        var task = _tasks.CreateTask(new TaskItem { Title = "Running" });
        _tasks.StartTimer(task);
        _tasks.SetUnbookedTicketSeconds(task, 7200, 0, 0);

        Assert.Equal(TaskStatus.Running, task.Status);
        using var connection = new SqliteConnection(_database.ConnectionString);
        connection.Open();
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT COUNT(*) FROM time_logs WHERE task_id=$id AND end_utc IS NULL";
        command.Parameters.AddWithValue("$id", task.Id.ToString());
        Assert.Equal(1L, Convert.ToInt64(command.ExecuteScalar()));
    }

    private static TicketTimeBooking Booking(Guid taskId) => new()
    {
        TaskId = taskId, TicketId = "42", TicketNumber = "20260042", BookingId = "b", ArticleId = "a",
        Minutes = 10, BookedMinutes = 10, SourceSeconds = 600, BookedAtUtc = DateTime.UtcNow,
        ShortDescription = "existing", Note = "unchanged", CostCenter = "c", Order = "o", Status = "Succeeded"
    };

    public void Dispose()
    {
        SqliteConnection.ClearAllPools();
        if (File.Exists(_databasePath)) File.Delete(_databasePath);
        if (File.Exists(_settingsPath)) File.Delete(_settingsPath);
        Assert.False(File.Exists(_databasePath));
    }
}
