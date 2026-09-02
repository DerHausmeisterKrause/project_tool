using System.Text.Json;
using TaskTool.Services;
using Xunit;

namespace TaskTool.Tests;

public sealed class SettingsServiceClientInstanceIdTests
{
    [Fact]
    public void MissingClientInstanceIdIsGeneratedPersistedAndStableAcrossRestart()
    {
        var directory = Path.Combine(Path.GetTempPath(), "plenaro-settings-test-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "settings.json");
        try
        {
            File.WriteAllText(path, "{}");
            var first = new SettingsService(new LoggerService(), path);
            var generated = first.Current.ClientInstanceId;

            Assert.True(Guid.TryParse(generated, out _));
            using (var persisted = JsonDocument.Parse(File.ReadAllText(path)))
                Assert.Equal(generated, persisted.RootElement.GetProperty("ClientInstanceId").GetString());

            var second = new SettingsService(new LoggerService(), path);
            Assert.Equal(generated, second.Current.ClientInstanceId);
        }
        finally { Directory.Delete(directory, recursive: true); }
    }
}
