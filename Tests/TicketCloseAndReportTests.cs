using System.Text.Json;
using Microsoft.Data.Sqlite;
using TaskTool.Models;
using TaskTool.Services;
using TaskTool.ViewModels;
using TaskTool.Views;
using Xunit;

namespace TaskTool.Tests;

public sealed class TicketCloseAndReportTests : IDisposable
{
    private readonly string _databasePath = Path.Combine(Path.GetTempPath(), $"plenaro-report-{Guid.NewGuid():N}.db");
    private readonly string _settingsPath = Path.Combine(Path.GetTempPath(), $"plenaro-report-settings-{Guid.NewGuid():N}.json");
    private readonly DatabaseService _database;
    private readonly TaskService _tasks;

    public TicketCloseAndReportTests()
    {
        var logger = new LoggerService();
        var settings = new SettingsService(logger, _settingsPath);
        settings.Current.OutlookSyncEnabled = false;
        _database = new DatabaseService(logger, _databasePath);
        _database.Initialize();
        _tasks = new TaskService(_database, logger, new OutlookInteropService(logger, settings), settings);
    }

    [Theory]
    [InlineData("Hallo Herr Müller,\n\nViele Grüße\nMarcel")]
    [InlineData("Hallo Herr Müller,\n\n\nViele Grüße\nMarcel")]
    public void CloseComposer_InsertsNoticeBelowGreetingAndPreservesSignature(string template)
    {
        var result = TicketCloseReplyComposer.EnsureCloseNotice(template);

        Assert.Equal($"Hallo Herr Müller,{Environment.NewLine}{Environment.NewLine}{TicketCloseReplyComposer.CloseNotice}{Environment.NewLine}{Environment.NewLine}Viele Grüße{Environment.NewLine}Marcel", result);
    }

    [Fact]
    public void CloseComposer_HandlesEmptyTemplateAndNeverDuplicatesNotice()
    {
        Assert.Equal(TicketCloseReplyComposer.CloseNotice, TicketCloseReplyComposer.EnsureCloseNotice(""));
        var once = TicketCloseReplyComposer.EnsureCloseNotice($"Hallo{Environment.NewLine}{Environment.NewLine}{TicketCloseReplyComposer.CloseNotice}");
        Assert.Equal(1, once.Split(TicketCloseReplyComposer.CloseNotice, StringSplitOptions.None).Length - 1);
    }

    [Fact]
    public void ClosePayload_ContainsArticleAndTicketStateWhileNormalReplyHasNoTicketObject()
    {
        var close = TicketSystemService.BuildTicketReplyPayload("session", "42", "customer@example.test", "Re: Test", "Body", TicketReplyMode.CloseTicket);
        var normal = TicketSystemService.BuildTicketReplyPayload("session", "42", "customer@example.test", "Re: Test", "Body", TicketReplyMode.Normal);

        using var closeJson = JsonDocument.Parse(JsonSerializer.Serialize(close));
        Assert.True(closeJson.RootElement.TryGetProperty("Article", out _));
        Assert.Equal("closed successful", closeJson.RootElement.GetProperty("Ticket").GetProperty("State").GetString());
        Assert.False(normal.ContainsKey("Ticket"));
    }

    [Fact]
    public void ManualAdjustmentsAreSummedDeletedAndDoNotAffectTodayBookings()
    {
        var month = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
        var first = _tasks.AddMonthlyReportTimeAdjustment(month, 2 * 3600 + 30 * 60, "Nachtrag");
        _tasks.AddMonthlyReportTimeAdjustment(month, 30 * 60);

        var adjustments = _tasks.GetMonthlyReportTimeAdjustments(month);
        Assert.Equal(3 * 3600, adjustments.Sum(item => item.Seconds));
        Assert.Equal(0, _tasks.GetSuccessfulBookingStatistics(DateTime.Today, TimeZoneInfo.Local).TodaySeconds);

        _tasks.DeleteMonthlyReportTimeAdjustment(first.Id);
        Assert.Equal(30 * 60, _tasks.GetMonthlyReportTimeAdjustments(month).Sum(item => item.Seconds));
    }

    [Fact]
    public void MonthlyReportAddsManualTimeToBookingsButKeepsTodayValueUnchanged()
    {
        var task = _tasks.CreateTask(new TaskItem { Title = "Ticket" });
        _tasks.CreateTicketTimeBooking(new TicketTimeBooking
        {
            TaskId = task.Id,
            TicketId = "42",
            TicketNumber = "20260042",
            Minutes = 6000,
            BookedMinutes = 6000,
            SourceSeconds = 360000,
            BookedAtUtc = DateTime.UtcNow,
            Status = "Succeeded"
        });
        var month = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
        _tasks.AddMonthlyReportTimeAdjustment(month, 2 * 3600 + 30 * 60);
        var logger = new LoggerService();
        var settings = new SettingsService(logger, _settingsPath);
        settings.Current.CalendarTimeZoneId = "UTC";

        var reports = new ReportsViewModel(_tasks, new WorkDayService(_database, logger), settings, new GermanTimeService(), logger);
        var current = Assert.Single(reports.MonthlyReports, item => item.IsCurrentMonth);

        Assert.Equal(360000, current.BookedSeconds);
        Assert.Equal(9000, current.ManualAdjustmentSeconds);
        Assert.Equal(369000, current.TotalSeconds);
        Assert.Equal(360000, reports.BookedSecondsToday);
    }

    [Fact]
    public void ManualAdjustmentMonthExistsWithoutTicketBookingsAndInvalidValuesAreRejected()
    {
        var month = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1).AddMonths(-2);
        _tasks.AddMonthlyReportTimeAdjustment(month, 60);

        Assert.Single(_tasks.GetMonthlyReportTimeAdjustments(month));
        Assert.Empty(_tasks.GetSuccessfulBookingStatistics(DateTime.Today, TimeZoneInfo.Local).SecondsByMonth);
        var logger = new LoggerService();
        var settings = new SettingsService(logger, _settingsPath);
        settings.Current.CalendarTimeZoneId = "UTC";
        var reports = new ReportsViewModel(_tasks, new WorkDayService(_database, logger), settings, new GermanTimeService(), logger);
        Assert.Contains(reports.MonthlyReports, item => item.Month == month && item.ManualAdjustmentSeconds == 60);
        Assert.Throws<ArgumentOutOfRangeException>(() => _tasks.AddMonthlyReportTimeAdjustment(month, 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => _tasks.AddMonthlyReportTimeAdjustment(month, -60));
        Assert.Throws<ArgumentOutOfRangeException>(() => _tasks.AddMonthlyReportTimeAdjustment(DateTime.Today.AddMonths(1), 60));
    }

    [Theory]
    [InlineData("0", "30", 1800, true)]
    [InlineData("2", "0", 7200, true)]
    [InlineData("0", "0", 0, false)]
    [InlineData("1", "75", 0, false)]
    [InlineData("-1", "0", 0, false)]
    public void AdjustmentInputValidatesHoursAndMinutes(string hours, string minutes, long expectedSeconds, bool expectedValid)
    {
        var valid = MonthlyReportTimeAdjustmentDialog.TryCalculateSeconds(hours, minutes, out var seconds, out _);
        Assert.Equal(expectedValid, valid);
        Assert.Equal(expectedSeconds, seconds);
    }

    [Fact]
    public void SchemaVersion31ContainsAdjustmentTableAndIndex()
    {
        using var connection = new SqliteConnection(_database.ConnectionString);
        connection.Open();
        using var command = connection.CreateCommand();
        command.CommandText = "SELECT version FROM schema_version";
        Assert.Equal(31L, command.ExecuteScalar());
        command.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='monthly_report_time_adjustments'";
        Assert.Equal(1L, command.ExecuteScalar());
        command.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type='index' AND name='idx_monthly_report_time_adjustments_month'";
        Assert.Equal(1L, command.ExecuteScalar());
    }

    public void Dispose()
    {
        SqliteConnection.ClearAllPools();
        if (File.Exists(_databasePath)) File.Delete(_databasePath);
        if (File.Exists(_settingsPath)) File.Delete(_settingsPath);
    }
}
