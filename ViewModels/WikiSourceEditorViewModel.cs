using TaskTool.Infrastructure;
using TaskTool.Models;

namespace TaskTool.ViewModels;

public sealed class WikiSourceEditorViewModel : ObservableObject
{
    private string _id = Guid.NewGuid().ToString(), _name = "Neue Wiki-Quelle", _providerType = "ConfluenceDataCenter", _baseUrl = string.Empty;
    private string _authMode = "BearerToken", _username = string.Empty, _secret = string.Empty, _spaceKey = string.Empty;
    private bool _enabled = true; private int _maxResults = 5;
    private bool _searchAllSpaces = true, _browserAutoSubmit = true, _isDefault;
    private string _spaceKeysText = string.Empty, _browserHomeUrl = string.Empty, _browserLoginMode = "BrowserSession", _browserUsername = string.Empty, _browserPassword = string.Empty;
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
    public bool SearchAllSpaces { get => _searchAllSpaces; set => Set(ref _searchAllSpaces, value); }
    public string SpaceKeysText { get => _spaceKeysText; set => Set(ref _spaceKeysText, value); }
    public string BrowserHomeUrl { get => _browserHomeUrl; set => Set(ref _browserHomeUrl, value); }
    public string BrowserLoginMode { get => _browserLoginMode; set { if (Set(ref _browserLoginMode, value)) { Raise(nameof(UsesWindowsIntegrated)); Raise(nameof(UsesBrowserCredentials)); } } }
    public string BrowserUsername { get => _browserUsername; set => Set(ref _browserUsername, value); }
    public string BrowserPassword { get => _browserPassword; set => Set(ref _browserPassword, value); }
    public bool BrowserAutoSubmit { get => _browserAutoSubmit; set => Set(ref _browserAutoSubmit, value); }
    public bool IsDefault { get => _isDefault; set => Set(ref _isDefault, value); }
    public bool UsesWindowsIntegrated => BrowserLoginMode == "WindowsIntegrated";
    public bool UsesBrowserCredentials => BrowserLoginMode == "UsernamePassword";
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
    public string BrowserPasswordEncrypted { get; private set; } = string.Empty;

    public static WikiSourceEditorViewModel FromModel(WikiSourceSettings source, string mask, bool isDefault = false) => new()
    { Id = source.Id, Name = source.Name, ProviderType = source.ProviderType, Enabled = source.Enabled, BaseUrl = source.BaseUrl, AuthMode = source.AuthMode, Username = source.Username, Secret = string.IsNullOrWhiteSpace(source.SecretEncrypted) ? string.Empty : mask, SecretEncrypted = source.SecretEncrypted, SpaceKey = source.SpaceKey, SearchAllSpaces = source.SearchAllSpaces, SpaceKeysText = string.Join(Environment.NewLine, source.SpaceKeys), BrowserHomeUrl = source.BrowserHomeUrl, BrowserLoginMode = source.BrowserLoginMode, BrowserUsername = source.BrowserUsername, BrowserPassword = string.IsNullOrWhiteSpace(source.BrowserPasswordEncrypted) ? string.Empty : mask, BrowserPasswordEncrypted = source.BrowserPasswordEncrypted, BrowserAutoSubmit = source.BrowserAutoSubmit, IsDefault = isDefault, MaxResults = source.MaxResults, ApiKeyHeaderName = source.ApiKeyHeaderName, HttpMethod = source.HttpMethod, SearchUrlTemplate = source.SearchUrlTemplate, RequestBodyTemplate = source.RequestBodyTemplate, ResultArrayPath = source.ResultArrayPath, ResultIdPath = source.ResultIdPath, ResultTitlePath = source.ResultTitlePath, ResultUrlPath = source.ResultUrlPath, ResultExcerptPath = source.ResultExcerptPath };

    public WikiSourceSettings ToModel() => new() { Id = Id.Trim(), Name = Name.Trim(), ProviderType = ProviderType, Enabled = Enabled, BaseUrl = BaseUrl.Trim().TrimEnd('/'), AuthMode = AuthMode, Username = Username.Trim(), SecretEncrypted = SecretEncrypted, ApiKeyHeaderName = ApiKeyHeaderName, SpaceKey = string.Empty, SearchAllSpaces = SearchAllSpaces, SpaceKeys = SpaceKeysText.Split(new[] { '\r', '\n', ',', ';' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).Distinct(StringComparer.OrdinalIgnoreCase).ToList(), BrowserHomeUrl = BrowserHomeUrl.Trim(), BrowserLoginMode = BrowserLoginMode, BrowserUsername = BrowserUsername.Trim(), BrowserPasswordEncrypted = BrowserPasswordEncrypted, BrowserAutoSubmit = BrowserAutoSubmit, MaxResults = MaxResults, HttpMethod = HttpMethod, SearchUrlTemplate = SearchUrlTemplate, RequestBodyTemplate = RequestBodyTemplate, ResultArrayPath = ResultArrayPath, ResultIdPath = ResultIdPath, ResultTitlePath = ResultTitlePath, ResultUrlPath = ResultUrlPath, ResultExcerptPath = ResultExcerptPath };
    public void SetEncryptedSecret(string value) => SecretEncrypted = value;
    public void SetEncryptedBrowserPassword(string value) => BrowserPasswordEncrypted = value;
}

public sealed record WikiChoice(string Value, string DisplayName);
