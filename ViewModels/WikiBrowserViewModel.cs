using System.Collections.ObjectModel;
using TaskTool.Infrastructure;
using TaskTool.Models;
using TaskTool.Services;

namespace TaskTool.ViewModels;

public sealed class WikiBrowserViewModel : ObservableObject
{
    private readonly SettingsService _settings;
    public string Title => "Wiki";
    public ObservableCollection<WikiSourceSettings> Sources { get; } = new();
    private WikiSourceSettings? _selectedSource;
    public WikiSourceSettings? SelectedSource { get => _selectedSource; set { if (Set(ref _selectedSource, value) && value != null) NavigationUrl = GetHomeUrl(value); } }
    private string _navigationUrl = string.Empty;
    public string NavigationUrl { get => _navigationUrl; private set => Set(ref _navigationUrl, value); }
    public WikiBrowserViewModel(SettingsService settings)
    {
        _settings = settings;
        RefreshSources();
    }

    public void RefreshSources()
    {
        var selectedId = SelectedSource?.Id; Sources.Clear();
        foreach (var source in _settings.Current.WikiSources.Where(x => x.Enabled && WikiSourceValidation.TryValidate(x, out _))) Sources.Add(source);
        _selectedSource = Sources.FirstOrDefault(x => x.Id == selectedId) ?? Sources.FirstOrDefault(x => x.Id == _settings.Current.DefaultWikiSourceId) ?? Sources.FirstOrDefault(); Raise(nameof(SelectedSource));
    }

    public void EnsureHome()
    {
        if (string.IsNullOrWhiteSpace(NavigationUrl) && SelectedSource != null) NavigationUrl = GetHomeUrl(SelectedSource);
    }

    public void NavigateTo(string sourceId, string url)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out var uri) || (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps)) return;
        SelectedSource = Sources.FirstOrDefault(x => x.Id == sourceId) ?? SelectedSource;
        NavigationUrl = uri.ToString();
    }

    public static string GetHomeUrl(WikiSourceSettings source)
    {
        if (!string.IsNullOrWhiteSpace(source.BrowserHomeUrl)) return source.BrowserHomeUrl;
        var baseUrl = source.BaseUrl.TrimEnd('/');
        return source.ProviderType == "ConfluenceCloud" && !baseUrl.EndsWith("/wiki", StringComparison.OrdinalIgnoreCase) ? baseUrl + "/wiki" : baseUrl;
    }
}
