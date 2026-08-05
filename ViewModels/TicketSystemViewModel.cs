using TaskTool.Infrastructure;
using TaskTool.Services;

namespace TaskTool.ViewModels;

public class TicketSystemViewModel : ObservableObject
{
    private readonly SettingsService _settings;

    public string Title => "Ticketsystem";
    public string TicketSystemWebUrl => _settings.Current.TicketSystemWebUrl;

    public TicketSystemViewModel(SettingsService settings)
    {
        _settings = settings;
    }

    public void Refresh()
    {
        Raise(nameof(TicketSystemWebUrl));
    }

    public override string ToString() => Title;
}
