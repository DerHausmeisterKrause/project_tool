using TaskTool.Models;
using TaskTool.Services;
using Xunit;

namespace TaskTool.Tests;

public sealed class TicketChangeNotificationTests
{
    [Fact]
    public void SettingDefaultsTrueAndSurvivesSaveAndReload()
    {
        var directory = Path.Combine(Path.GetTempPath(), "plenaro-change-settings-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "settings.json");
        try
        {
            File.WriteAllText(path, "{}");
            var settings = new SettingsService(new LoggerService(), path);
            Assert.True(settings.Current.NotifyOnTicketChanges);

            settings.Current.NotifyOnTicketChanges = false;
            settings.Save();
            Assert.False(new SettingsService(new LoggerService(), path).Current.NotifyOnTicketChanges);
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    [Fact]
    public void IndividualPayloadsContainTaskTicketNumberAndTitleUpToSharedLimit()
    {
        var taskId = Guid.NewGuid();
        var candidates = Enumerable.Range(1, 5).ToDictionary(
            index => index.ToString(),
            index => (index == 1 ? taskId : Guid.NewGuid(), $"20260000{index}", index == 1 ? "Drucker defekt" : string.Empty));

        var payloads = TicketSystemService.BuildTicketChangeNotificationPayloads(candidates);

        Assert.Equal(5, payloads.Count);
        Assert.Contains(payloads, payload => payload.TaskId == taskId
            && payload.Text == "Neue Nachricht in Ticket 202600001\nDrucker defekt");
        Assert.Contains(payloads, payload => payload.Text == "Neue Nachricht in Ticket 202600002");
    }

    [Fact]
    public void MoreThanSharedLimitCreatesOneSummaryPayload()
    {
        var candidates = Enumerable.Range(1, 6).ToDictionary(
            index => index.ToString(), index => (Guid.NewGuid(), $"20260000{index}", "Titel"));

        var payload = Assert.Single(TicketSystemService.BuildTicketChangeNotificationPayloads(candidates));

        Assert.Equal(Guid.Empty, payload.TaskId);
        Assert.Equal("Neue Nachrichten in 6 Tickets\nÖffne Plenaro, um die Änderungen anzusehen.", payload.Text);
    }
}
