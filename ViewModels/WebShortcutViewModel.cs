using TaskTool.Models;
using TaskTool.Services;

namespace TaskTool.ViewModels;

public sealed class WebShortcutViewModel
{
    public WebShortcutSettings Shortcut { get; } public string ShortcutId=>Shortcut.Id; public string Url=>Shortcut.Url;
    public string Title=>!string.IsNullOrWhiteSpace(Shortcut.Name)?Shortcut.Name:Uri.TryCreate(Shortcut.Url,UriKind.Absolute,out var uri)?uri.Host:"Webseite";
    public WebShortcutViewModel(WebShortcutSettings shortcut)=>Shortcut=shortcut;
}
