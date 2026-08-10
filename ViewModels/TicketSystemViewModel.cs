using TaskTool.Infrastructure;
using TaskTool.Services;

namespace TaskTool.ViewModels;

public class TicketSystemViewModel : ObservableObject
{
    private readonly SettingsService _settings;
    private string? _navigationTargetUrl;

    public string Title => "Ticketsystem";
    public string TicketSystemWebUrl => _settings.Current.TicketSystemWebUrl;
    public string NavigationUrl => _navigationTargetUrl ?? TicketSystemWebUrl;
    public bool AutofillCredentials => _settings.Current.TicketSystemAutofillCredentials;
    public bool AutoLogin => _settings.Current.TicketSystemAutoLogin;
    public string Username => _settings.Current.TicketSystemUsername;
    public string Password => _settings.GetTicketSystemPassword();

    public TicketSystemViewModel(SettingsService settings)
    {
        _settings = settings;
    }

    public void Refresh()
    {
        Raise(nameof(TicketSystemWebUrl));
        Raise(nameof(NavigationUrl));
    }

    public void NavigateTo(string url)
    {
        _navigationTargetUrl = url;
        Raise(nameof(NavigationUrl));
    }

    public override string ToString() => Title;
}
