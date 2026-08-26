using System.Windows;

namespace TaskTool.Views;

public partial class UpdateAvailableDialog : Window
{
    public UpdateAvailableDialog(string installedVersion, string availableVersion, string releaseNotes)
    {
        InitializeComponent();
        DataContext = new UpdateAvailableDialogModel(installedVersion, availableVersion, TrimReleaseNotes(releaseNotes));
    }

    private void Install_Click(object sender, RoutedEventArgs e) => DialogResult = true;
    private void Defer_Click(object sender, RoutedEventArgs e) => DialogResult = false;

    private static string TrimReleaseNotes(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return string.Empty;
        var trimmed = value.Trim();
        return trimmed.Length <= 1200 ? trimmed : trimmed[..1200].TrimEnd() + "…";
    }

    private sealed record UpdateAvailableDialogModel(string InstalledVersion, string AvailableVersion, string ReleaseNotes)
    {
        public string IntroText => $"Plenaro {AvailableVersion} ist verfügbar.";
        public bool HasReleaseNotes => !string.IsNullOrWhiteSpace(ReleaseNotes);
    }
}
