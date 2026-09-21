using System.Windows;

namespace TaskTool.Services;

public interface IClipboardService
{
    void SetText(string text);
}

public sealed class ClipboardService : IClipboardService
{
    public void SetText(string text) => Clipboard.SetText(text);
}
