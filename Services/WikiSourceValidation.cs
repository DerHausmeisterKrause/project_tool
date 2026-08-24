using TaskTool.Models;

namespace TaskTool.Services;

public static class WikiSourceValidation
{
    public static IReadOnlyList<string> ProviderTypes { get; } = new[] { "ConfluenceDataCenter", "ConfluenceCloud", "GenericRest", "XWiki" };
    public static IReadOnlyList<string> AuthModes { get; } = new[] { "BearerToken", "UsernameToken", "Basic", "ApiKey" };

    public static bool TryValidate(WikiSourceSettings? source, out string error)
    {
        if (source == null) { error = "Bitte zuerst eine Wiki-Quelle auswählen."; return false; }
        if (string.IsNullOrWhiteSpace(source.Id)) { error = "Wiki-Quelle besitzt keine gültige ID."; return false; }
        if (string.IsNullOrWhiteSpace(source.Name)) { error = "Bitte einen Namen für die Wiki-Quelle eingeben."; return false; }
        if (string.IsNullOrWhiteSpace(source.ProviderType) || !ProviderTypes.Contains(source.ProviderType, StringComparer.OrdinalIgnoreCase)) { error = "Bitte einen gültigen Wiki-Typ auswählen."; return false; }
        if (string.IsNullOrWhiteSpace(source.BaseUrl) || !Uri.TryCreate(source.BaseUrl, UriKind.Absolute, out var uri) || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps)) { error = "Bitte eine gültige absolute Base URL eingeben, z.B. https://wiki.example.de."; return false; }
        if (string.IsNullOrWhiteSpace(source.AuthMode) || !AuthModes.Contains(source.AuthMode, StringComparer.OrdinalIgnoreCase)) { error = "Bitte eine gültige Authentifizierung auswählen."; return false; }
        if (source.MaxResults is < 1 or > 20) { error = "Max. Ergebnisse muss zwischen 1 und 20 liegen."; return false; }
        error = string.Empty; return true;
    }
}
