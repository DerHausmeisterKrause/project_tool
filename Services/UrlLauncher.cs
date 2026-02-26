using System.Diagnostics;

namespace TaskTool.Services;

public static class UrlLauncher
{
    public static bool TryOpen(string? url, out string error)
    {
        error = string.Empty;

        if (string.IsNullOrWhiteSpace(url))
        {
            error = "Ticket URL ist leer.";
            return false;
        }

        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps))
        {
            error = "Ticket URL ist ungültig (nur http/https).";
            return false;
        }

        try
        {
            Process.Start(new ProcessStartInfo(uri.AbsoluteUri) { UseShellExecute = true });
            return true;
        }
        catch (Exception ex)
        {
            error = $"Ticket URL konnte nicht geöffnet werden: {ex.Message}";
            return false;
        }
    }
}
