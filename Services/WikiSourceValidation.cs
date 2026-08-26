using TaskTool.Models;

namespace TaskTool.Services;

public static class WikiSourceValidation
{
    public static IReadOnlyList<string> ProviderTypes { get; } = new[] { "ConfluenceDataCenter", "ConfluenceCloud", "GenericRest", "XWiki" };
    public static IReadOnlyList<string> AuthModes { get; } = new[] { "BearerToken", "UsernameToken", "Basic", "ApiKey", "WindowsIntegrated" };

    public static bool TryValidate(WikiSourceSettings? source, out string error)
    {
        if (source == null) { error = "Bitte zuerst eine Wiki-Quelle auswählen."; return false; }
        if (string.IsNullOrWhiteSpace(source.Id)) { error = "Wiki-Quelle besitzt keine gültige ID."; return false; }
        if (string.IsNullOrWhiteSpace(source.Name)) { error = "Bitte einen Namen für die Wiki-Quelle eingeben."; return false; }
        if (string.IsNullOrWhiteSpace(source.ProviderType) || !ProviderTypes.Contains(source.ProviderType, StringComparer.OrdinalIgnoreCase)) { error = "Bitte einen gültigen Wiki-Typ auswählen."; return false; }
        if (string.IsNullOrWhiteSpace(source.BaseUrl) || !Uri.TryCreate(source.BaseUrl, UriKind.Absolute, out var uri) || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps)) { error = "Bitte eine gültige absolute Base URL eingeben, z.B. https://wiki.example.de."; return false; }
        if (string.IsNullOrWhiteSpace(source.AuthMode) || !AuthModes.Contains(source.AuthMode, StringComparer.OrdinalIgnoreCase)) { error = "Bitte eine gültige Authentifizierung auswählen."; return false; }
        if (source.MaxResults is < 1 or > 20) { error = "Max. Ergebnisse muss zwischen 1 und 20 liegen."; return false; }
        if (!source.SearchAllSpaces && source.ProviderType.StartsWith("Confluence", StringComparison.OrdinalIgnoreCase) && (source.SpaceKeys == null || source.SpaceKeys.Count == 0)) { error = "Bitte mindestens einen Space Key eingeben oder 'Alle Bereiche durchsuchen' aktivieren."; return false; }
        if (!string.IsNullOrWhiteSpace(source.BrowserHomeUrl) && (!Uri.TryCreate(source.BrowserHomeUrl, UriKind.Absolute, out var homeUri) || (homeUri.Scheme != Uri.UriSchemeHttp && homeUri.Scheme != Uri.UriSchemeHttps))) { error = "Bitte eine gültige Browser-Startseite mit http oder https eingeben."; return false; }
        if (source.BrowserLoginMode == "UsernamePassword" && uri.Scheme != Uri.UriSchemeHttps) { error = "Browser-Autologin ist nur mit einer HTTPS Base URL erlaubt."; return false; }
        if (source.BrowserLoginMode == "UsernamePassword" && (string.IsNullOrWhiteSpace(source.BrowserUsername) || string.IsNullOrWhiteSpace(source.BrowserPasswordEncrypted))) { error = "Bitte Browser-Benutzername und Browser-Passwort eingeben."; return false; }
        error = string.Empty; return true;
    }
}
