using TaskTool.Infrastructure;
using TaskTool.Models;

namespace TaskTool.ViewModels;

public sealed class WikiSourceEditorViewModel : ObservableObject
{
    private string _id = Guid.NewGuid().ToString(), _name = "Neue Wiki-Quelle", _providerType = "ConfluenceDataCenter", _baseUrl = string.Empty;
    private string _authMode = "BearerToken", _username = string.Empty, _secret = string.Empty, _spaceKey = string.Empty;
    private bool _enabled = true; private int _maxResults = 5;
    public string Id { get => _id; set => Set(ref _id, value); }
    public string Name { get => _name; set => Set(ref _name, value); }
    public string ProviderType { get => _providerType; set { if (!string.IsNullOrWhiteSpace(value)) Set(ref _providerType, value); } }
    public bool Enabled { get => _enabled; set => Set(ref _enabled, value); }
    public string BaseUrl { get => _baseUrl; set => Set(ref _baseUrl, value); }
    public string AuthMode { get => _authMode; set { if (!string.IsNullOrWhiteSpace(value)) Set(ref _authMode, value); } }
    public string Username { get => _username; set => Set(ref _username, value); }
    public string Secret { get => _secret; set => Set(ref _secret, value); }
    public string SpaceKey { get => _spaceKey; set => Set(ref _spaceKey, value); }
    public int MaxResults { get => _maxResults; set => Set(ref _maxResults, value); }
    public string SecretEncrypted { get; private set; } = string.Empty;
    public string ApiKeyHeaderName { get; set; } = "X-API-Key";
    public string HttpMethod { get; set; } = "GET";
    public string SearchUrlTemplate { get; set; } = string.Empty;
    public string RequestBodyTemplate { get; set; } = string.Empty;
    public string ResultArrayPath { get; set; } = "$.results";
    public string ResultIdPath { get; set; } = "id";
    public string ResultTitlePath { get; set; } = "title";
    public string ResultUrlPath { get; set; } = "url";
    public string ResultExcerptPath { get; set; } = "excerpt";

    public static WikiSourceEditorViewModel FromModel(WikiSourceSettings source, string mask) => new()
    { Id = source.Id, Name = source.Name, ProviderType = source.ProviderType, Enabled = source.Enabled, BaseUrl = source.BaseUrl, AuthMode = source.AuthMode, Username = source.Username, Secret = string.IsNullOrWhiteSpace(source.SecretEncrypted) ? string.Empty : mask, SecretEncrypted = source.SecretEncrypted, SpaceKey = source.SpaceKey, MaxResults = source.MaxResults, ApiKeyHeaderName = source.ApiKeyHeaderName, HttpMethod = source.HttpMethod, SearchUrlTemplate = source.SearchUrlTemplate, RequestBodyTemplate = source.RequestBodyTemplate, ResultArrayPath = source.ResultArrayPath, ResultIdPath = source.ResultIdPath, ResultTitlePath = source.ResultTitlePath, ResultUrlPath = source.ResultUrlPath, ResultExcerptPath = source.ResultExcerptPath };

    public WikiSourceSettings ToModel() => new() { Id = Id.Trim(), Name = Name.Trim(), ProviderType = ProviderType, Enabled = Enabled, BaseUrl = BaseUrl.Trim().TrimEnd('/'), AuthMode = AuthMode, Username = Username.Trim(), SecretEncrypted = SecretEncrypted, ApiKeyHeaderName = ApiKeyHeaderName, SpaceKey = SpaceKey.Trim(), MaxResults = MaxResults, HttpMethod = HttpMethod, SearchUrlTemplate = SearchUrlTemplate, RequestBodyTemplate = RequestBodyTemplate, ResultArrayPath = ResultArrayPath, ResultIdPath = ResultIdPath, ResultTitlePath = ResultTitlePath, ResultUrlPath = ResultUrlPath, ResultExcerptPath = ResultExcerptPath };
    public void SetEncryptedSecret(string value) => SecretEncrypted = value;
}

public sealed record WikiChoice(string Value, string DisplayName);
