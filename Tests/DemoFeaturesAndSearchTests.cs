using System.Security.Cryptography;
using TaskTool.Infrastructure;
using TaskTool.Models;
using TaskTool.Services;
using TaskTool.ViewModels;

namespace TaskTool.Tests;

public sealed class DemoFeaturesAndSearchTests
{
    [Fact]
    public void FreshSettings_AreLocked()
    {
        using var directory = new TemporaryDirectory();
        var settings = new SettingsService(new LoggerService(), Path.Combine(directory.Path, "settings.json"));

        Assert.Empty(settings.Current.DemoFeaturesUnlockGrantEncrypted);
        Assert.False(settings.IsDemoFeaturesUnlocked);
    }

    [Fact]
    public void WrongPassword_DoesNotCreateGrant()
    {
        using var directory = new TemporaryDirectory();
        var settings = new SettingsService(new LoggerService(), Path.Combine(directory.Path, "settings.json"));

        Assert.False(settings.TryUnlockDemoFeatures("definitely-not-the-production-secret"));
        Assert.Empty(settings.Current.DemoFeaturesUnlockGrantEncrypted);
    }

    [Fact]
    public void Verifier_AcceptsMatchingTestHashAndRejectsWrongValue()
    {
        const string testSecret = "unit-test-only-secret";
        const int iterations = 200_000;
        var salt = Enumerable.Range(1, 16).Select(value => (byte)value).ToArray();
        var hash = Rfc2898DeriveBytes.Pbkdf2(testSecret, salt, iterations, HashAlgorithmName.SHA256, 32);

        Assert.True(DemoFeatureUnlockVerifier.Verify(testSecret, Convert.ToBase64String(salt), Convert.ToBase64String(hash), iterations));
        Assert.False(DemoFeatureUnlockVerifier.Verify("wrong", Convert.ToBase64String(salt), Convert.ToBase64String(hash), iterations));
    }

    [Fact]
    public void ArbitraryGrantString_DoesNotUnlock()
    {
        using var directory = new TemporaryDirectory();
        var settings = new SettingsService(new LoggerService(), Path.Combine(directory.Path, "settings.json"));
        settings.Current.DemoFeaturesUnlockGrantEncrypted = Convert.ToBase64String("arbitrary"u8);

        Assert.False(settings.IsDemoFeaturesUnlocked);
    }

    [Fact]
    public async Task DisabledAvailability_DoesNotCallAnyLoader()
    {
        var outlookCalls = 0;
        var ticketCalls = 0;
        var service = new AgentAvailabilityService(
            _ => { outlookCalls++; return Task.FromResult<OutlookFreeBusySnapshot?>(null); },
            _ => { ticketCalls++; return Task.FromResult<int?>(0); },
            featureEnabled: () => false);

        var result = await service.GetAvailabilityAsync(new ZnunyAgent { UserId = 1 });

        Assert.Equal(AgentAvailabilityLevel.Unknown, result.Level);
        Assert.Equal(0, outlookCalls);
        Assert.Equal(0, ticketCalls);
    }

    [Fact]
    public void AvailabilityResultCannotApplyAfterFeatureWasLocked()
    {
        var agent = new ZnunyAgent { UserId = 1 };
        Assert.False(TodayViewModel.CanApplyAvailabilityResult(agent, agent, featureEnabled: false));
        Assert.True(TodayViewModel.CanApplyAvailabilityResult(agent, agent, featureEnabled: true));
    }

    [Theory]
    [InlineData("4711", true)]
    [InlineData("infra", true)]
    [InlineData("ROSENHEIM", true)]
    [InlineData(" it ", true)]
    [InlineData("foo", false)]
    [InlineData("", true)]
    public void Search_MatchesKeyOrDisplayTextWithoutChangingOption(string query, bool expected)
    {
        var option = new TicketFieldOption("4711", "IT Infrastruktur Rosenheim");
        var original = option;

        Assert.Equal(expected, SearchableComboBoxAssist.Matches(option, query));
        Assert.Same(original, option);
        Assert.Equal("4711", option.Key);
    }

    [Fact]
    public void TaskDetail_UsesSearchAssistWithoutFreeValueBindings()
    {
        var xaml = File.ReadAllText(Path.Combine(RepoRoot(), "Views", "TaskDetailPanel.xaml"));
        Assert.Equal(2, Count(xaml, "SearchableComboBoxAssist.IsEnabled=\"True\""));
        Assert.Contains("SelectedItem=\"{Binding SelectedCostCenter}\"", xaml, StringComparison.Ordinal);
        Assert.Contains("SelectedItem=\"{Binding SelectedOrder}\"", xaml, StringComparison.Ordinal);
        Assert.DoesNotContain("Text=\"{Binding SelectedCostCenter", xaml, StringComparison.Ordinal);
        Assert.DoesNotContain("Text=\"{Binding SelectedOrder", xaml, StringComparison.Ordinal);
    }

    private static int Count(string value, string part)
        => (value.Length - value.Replace(part, string.Empty, StringComparison.Ordinal).Length) / part.Length;

    private static string RepoRoot()
        => Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));

    private sealed class TemporaryDirectory : IDisposable
    {
        public string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "plenaro-demo-tests-" + Guid.NewGuid().ToString("N"));
        public TemporaryDirectory() => Directory.CreateDirectory(Path);
        public void Dispose() { try { Directory.Delete(Path, true); } catch { } }
    }
}
