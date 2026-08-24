using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using TaskTool.Models;

namespace TaskTool.Services;

public interface IWikiProvider
{
    string ProviderType { get; }
    Task<IReadOnlyList<WikiProviderResult>> SearchAsync(WikiSourceSettings source, IReadOnlyList<string> terms, int limit, CancellationToken cancellationToken);
}

public abstract class HttpWikiProvider(SettingsService settings) : IWikiProvider
{
    protected readonly SettingsService Settings = settings;
    public abstract string ProviderType { get; }
    protected HttpClient CreateClient(WikiSourceSettings source)
    {
        var client = new HttpClient(new HttpClientHandler { AllowAutoRedirect = false }) { Timeout = TimeSpan.FromSeconds(10) };
        var secret = Settings.GetWikiSecret(source);
        if (source.AuthMode.Equals("BearerToken", StringComparison.OrdinalIgnoreCase))
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", secret);
        else if (source.AuthMode.Equals("Basic", StringComparison.OrdinalIgnoreCase) || source.AuthMode.Equals("UsernameToken", StringComparison.OrdinalIgnoreCase))
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Convert.ToBase64String(Encoding.UTF8.GetBytes($"{source.Username}:{secret}")));
        else if (source.AuthMode.Equals("ApiKey", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(source.ApiKeyHeaderName))
            client.DefaultRequestHeaders.TryAddWithoutValidation(source.ApiKeyHeaderName, secret);
        return client;
    }
    protected static string Clean(string? value, bool stripHtml = false)
    {
        value = (value ?? string.Empty).Replace("@@@hl@@@", "").Replace("@@@endhl@@@", "");
        if (stripHtml) value = Regex.Replace(value, "<[^>]+>", " ");
        return WebUtility.HtmlDecode(Regex.Replace(value, @"\s+", " ")).Trim();
    }
    public abstract Task<IReadOnlyList<WikiProviderResult>> SearchAsync(WikiSourceSettings source, IReadOnlyList<string> terms, int limit, CancellationToken cancellationToken);
}

public class ConfluenceDataCenterWikiProvider(SettingsService settings) : HttpWikiProvider(settings)
{
    public override string ProviderType => "ConfluenceDataCenter";
    protected virtual string ApiPath => "/rest/api/search";
    public override async Task<IReadOnlyList<WikiProviderResult>> SearchAsync(WikiSourceSettings source, IReadOnlyList<string> terms, int limit, CancellationToken token)
    {
        var cqlTerms = string.Join(" OR ", terms.Select(t => $"siteSearch~'{EscapeCql(t)}'"));
        var cql = $"type=page AND ({cqlTerms})" + (string.IsNullOrWhiteSpace(source.SpaceKey) ? "" : $" AND space='{EscapeCql(source.SpaceKey)}'");
        var endpoint = source.BaseUrl.TrimEnd('/') + ApiPath + "?cql=" + Uri.EscapeDataString(cql) + "&limit=" + limit;
        using var client = CreateClient(source);
        using var response = await client.GetAsync(endpoint, token);
        response.EnsureSuccessStatusCode();
        using var doc = JsonDocument.Parse(await response.Content.ReadAsStringAsync(token));
        var root = doc.RootElement;
        var baseUrl = root.TryGetProperty("_links", out var links) && links.TryGetProperty("base", out var b) ? b.GetString() : source.BaseUrl;
        var list = new List<WikiProviderResult>(); var rank = 0;
        foreach (var item in root.GetProperty("results").EnumerateArray())
        {
            var content = item.TryGetProperty("content", out var c) ? c : default;
            var title = content.ValueKind == JsonValueKind.Object && content.TryGetProperty("title", out var ct) ? ct.GetString() : item.TryGetProperty("title", out var t) ? t.GetString() : "";
            var id = content.ValueKind == JsonValueKind.Object && content.TryGetProperty("id", out var ci) ? ci.GetString() : "";
            var url = item.TryGetProperty("url", out var u) ? u.GetString() ?? "" : "";
            if (Uri.TryCreate(url, UriKind.Relative, out _) && Uri.TryCreate((baseUrl ?? source.BaseUrl).TrimEnd('/') + "/", UriKind.Absolute, out var baseUri)) url = new Uri(baseUri, url.TrimStart('/')).ToString();
            list.Add(new(id ?? "", Clean(title), url, Clean(item.TryGetProperty("excerpt", out var e) ? e.GetString() : "", true), rank++));
        }
        return list;
    }
    public static string EscapeCql(string value) => value.Replace("\\", "\\\\").Replace("'", "\\'").Replace("\r", " ").Replace("\n", " ");
}

public sealed class ConfluenceCloudWikiProvider(SettingsService settings) : ConfluenceDataCenterWikiProvider(settings)
{
    public override string ProviderType => "ConfluenceCloud";
    protected override string ApiPath => "/wiki/rest/api/search";
    public override Task<IReadOnlyList<WikiProviderResult>> SearchAsync(WikiSourceSettings source, IReadOnlyList<string> terms, int limit, CancellationToken token)
    {
        source.BaseUrl = Regex.Replace(source.BaseUrl.TrimEnd('/'), "/wiki$", "", RegexOptions.IgnoreCase);
        return base.SearchAsync(source, terms, limit, token);
    }
}

public sealed class GenericRestWikiProvider(SettingsService settings) : HttpWikiProvider(settings)
{
    public override string ProviderType => "GenericRest";
    public override async Task<IReadOnlyList<WikiProviderResult>> SearchAsync(WikiSourceSettings source, IReadOnlyList<string> terms, int limit, CancellationToken token)
    {
        var query = string.Join(" ", terms); var queryJson = JsonSerializer.Serialize(query)[1..^1];
        string Expand(string s) => s.Replace("{queryUrl}", Uri.EscapeDataString(query)).Replace("{queryJson}", queryJson).Replace("{limit}", limit.ToString());
        var url = Expand(string.IsNullOrWhiteSpace(source.SearchUrlTemplate) ? source.BaseUrl : source.SearchUrlTemplate);
        if (Uri.TryCreate(url, UriKind.Relative, out _) && Uri.TryCreate(source.BaseUrl.TrimEnd('/') + "/", UriKind.Absolute, out var root)) url = new Uri(root, url).ToString();
        using var client = CreateClient(source); using var request = new HttpRequestMessage(source.HttpMethod.Equals("POST", StringComparison.OrdinalIgnoreCase) ? HttpMethod.Post : HttpMethod.Get, url);
        if (request.Method == HttpMethod.Post) request.Content = new StringContent(Expand(source.RequestBodyTemplate), Encoding.UTF8, "application/json");
        using var response = await client.SendAsync(request, token); response.EnsureSuccessStatusCode();
        using var doc = JsonDocument.Parse(await response.Content.ReadAsStringAsync(token));
        var array = Resolve(doc.RootElement, source.ResultArrayPath.Replace("$.", "")); var results = new List<WikiProviderResult>(); var rank = 0;
        foreach (var item in array.EnumerateArray()) results.Add(new(Get(item, source.ResultIdPath), Clean(Get(item, source.ResultTitlePath)), Get(item, source.ResultUrlPath), Clean(Get(item, source.ResultExcerptPath), true), rank++));
        return results;
    }
    private static JsonElement Resolve(JsonElement item, string path) { foreach (var part in path.Split('.', StringSplitOptions.RemoveEmptyEntries)) item = item.GetProperty(part); return item; }
    private static string Get(JsonElement item, string path) { try { var value = Resolve(item, path); return value.ValueKind == JsonValueKind.String ? value.GetString() ?? "" : value.ToString(); } catch { return ""; } }
}

public sealed class XWikiProvider(SettingsService settings) : HttpWikiProvider(settings)
{
    public override string ProviderType => "XWiki";
    public override async Task<IReadOnlyList<WikiProviderResult>> SearchAsync(WikiSourceSettings source, IReadOnlyList<string> terms, int limit, CancellationToken token)
    {
        using var client = CreateClient(source); var url = source.BaseUrl.TrimEnd('/') + "/rest/wikis/xwiki/search?q=" + Uri.EscapeDataString(string.Join(" ", terms)) + "&number=" + limit;
        var xml = await client.GetStringAsync(url, token); var doc = XDocument.Parse(xml); var rank = 0;
        return doc.Descendants().Where(x => x.Name.LocalName == "searchResult").Select(x => new WikiProviderResult(x.Elements().FirstOrDefault(e => e.Name.LocalName == "id")?.Value ?? "", Clean(x.Elements().FirstOrDefault(e => e.Name.LocalName == "title")?.Value), x.Elements().FirstOrDefault(e => e.Name.LocalName == "url")?.Value ?? "", "", rank++)).ToList();
    }
}
