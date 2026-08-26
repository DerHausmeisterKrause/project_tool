using TaskTool.Infrastructure;
using TaskTool.Models;

namespace TaskTool.ViewModels;

public sealed class WebShortcutViewModel : ObservableObject
{
    public WebShortcutSettings Shortcut { get; private set; }
    public string ShortcutId => Shortcut.Id;
    public string Url => Shortcut.Url;
    public string EnvironmentKey => $"{ShortcutId}|{(Shortcut.DisableWebSecurity ? "cors-disabled" : "normal")}";
    public string Title => !string.IsNullOrWhiteSpace(Shortcut.Name)
        ? Shortcut.Name
        : Uri.TryCreate(Shortcut.Url, UriKind.Absolute, out var uri) ? uri.Host : "Webseite";

    public event Action? BrowserConfigurationChanged;

    public WebShortcutViewModel(WebShortcutSettings shortcut) => Shortcut = shortcut;

    public void Update(WebShortcutSettings shortcut)
    {
        if (!string.Equals(ShortcutId, shortcut.Id, StringComparison.Ordinal))
            throw new ArgumentException("Die Shortcut-ID darf beim Aktualisieren nicht geändert werden.", nameof(shortcut));

        var browserConfigurationChanged = !string.Equals(Shortcut.Url, shortcut.Url, StringComparison.Ordinal)
            || Shortcut.DisableWebSecurity != shortcut.DisableWebSecurity;
        Shortcut = shortcut;
        Raise(nameof(Title));
        Raise(nameof(Url));
        Raise(nameof(EnvironmentKey));
        if (browserConfigurationChanged) BrowserConfigurationChanged?.Invoke();
    }
}
