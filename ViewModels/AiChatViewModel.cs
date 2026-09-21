using System.Collections.ObjectModel;
using System.Net;
using System.Net.Http;
using TaskTool.Infrastructure;
using TaskTool.Models;
using TaskTool.Services;

namespace TaskTool.ViewModels;

public sealed class AiChatViewModel : ObservableObject
{
    public const int MaxContextMessages = 20;
    public const string SystemPrompt = """
        Du bist der KI-Assistent in Plenaro.
        Antworte hilfreich, sachlich und präzise.
        Antworte standardmäßig auf Deutsch, sofern der Benutzer keine andere Sprache verwendet.
        Erfinde keine angeblichen Plenaro-, Znuny- oder Wiki-Daten, die dir nicht übergeben wurden.
        """;

    private readonly IAiChatService _ai;
    private readonly IClipboardService _clipboard;
    private readonly Action _openAiSettings;
    private string _inputText = string.Empty;
    private string _errorMessage = string.Empty;
    private bool _isSending;

    public AiChatViewModel(IAiChatService ai, IClipboardService clipboard, Action openAiSettings)
    {
        _ai = ai;
        _clipboard = clipboard;
        _openAiSettings = openAiSettings;
        SendCommand = new RelayCommand(async () => await SendAsync(), () => CanSend);
        ClearCommand = new RelayCommand(Clear, () => Messages.Count > 0 && !IsSending);
        CopyMessageCommand = new RelayCommand<AiChatMessage>(CopyMessage, CanCopyMessage);
        OpenAiSettingsCommand = new RelayCommand(_openAiSettings);
        _ai.StateChanged += (_, _) => RefreshState();
    }

    public string Title => "KI";
    public ObservableCollection<AiChatMessage> Messages { get; } = new();
    public RelayCommand SendCommand { get; }
    public RelayCommand ClearCommand { get; }
    public RelayCommand<AiChatMessage> CopyMessageCommand { get; }
    public RelayCommand OpenAiSettingsCommand { get; }
    public bool IsEnabled => _ai.IsEnabled;
    public bool IsDisabled => !IsEnabled;
    public bool CanSend => !IsSending && _ai.CanChat && !string.IsNullOrWhiteSpace(InputText);
    public string ProviderDescription => _ai.ProviderDescription;
    public string AvailabilityMessage => _ai.AvailabilityMessage;
    public string SendingStatus => IsSending ? "KI antwortet …" : string.Empty;
    public bool HasMessages => Messages.Count > 0;

    public string InputText
    {
        get => _inputText;
        set { if (Set(ref _inputText, value)) SendCommand.RaiseCanExecuteChanged(); }
    }

    public string ErrorMessage
    {
        get => _errorMessage;
        private set => Set(ref _errorMessage, value);
    }

    public bool IsSending
    {
        get => _isSending;
        private set
        {
            if (!Set(ref _isSending, value)) return;
            Raise(nameof(CanSend));
            Raise(nameof(SendingStatus));
            SendCommand.RaiseCanExecuteChanged();
            ClearCommand.RaiseCanExecuteChanged();
        }
    }

    public async Task SendAsync(CancellationToken cancellationToken = default)
    {
        if (!CanSend) return;
        var text = InputText.Trim();
        Messages.Add(new AiChatMessage(AiChatRole.User, text, DateTime.Now));
        Raise(nameof(HasMessages));
        InputText = string.Empty;
        ErrorMessage = string.Empty;
        IsSending = true;
        try
        {
            var answer = await _ai.ChatAsync(BuildRequestMessages(), new AiRequestOptions(0.3, 1024), cancellationToken);
            Messages.Add(new AiChatMessage(AiChatRole.Assistant, answer, DateTime.Now));
        }
        catch (Exception exception)
        {
            ErrorMessage = $"Die KI-Anfrage ist fehlgeschlagen: {DescribeError(exception)}";
        }
        finally
        {
            IsSending = false;
            ClearCommand.RaiseCanExecuteChanged();
        }
    }

    public IReadOnlyList<AiChatRequestMessage> BuildRequestMessages()
    {
        var context = Messages.TakeLast(MaxContextMessages)
            .Select(message => new AiChatRequestMessage(message.Role, message.Content));
        return new[] { new AiChatRequestMessage(AiChatRole.System, SystemPrompt) }.Concat(context).ToArray();
    }

    private void Clear()
    {
        Messages.Clear();
        Raise(nameof(HasMessages));
        ErrorMessage = string.Empty;
        ClearCommand.RaiseCanExecuteChanged();
    }

    public void ClearChat() => Clear();

    private bool CanCopyMessage(AiChatMessage? message)
        => message?.IsAssistant == true && !string.IsNullOrWhiteSpace(message.Content);

    private void CopyMessage(AiChatMessage? message)
    {
        if (CanCopyMessage(message)) _clipboard.SetText(message!.Content);
    }

    private void RefreshState()
    {
        Raise(nameof(IsEnabled)); Raise(nameof(IsDisabled)); Raise(nameof(CanSend));
        Raise(nameof(ProviderDescription)); Raise(nameof(AvailabilityMessage));
        SendCommand.RaiseCanExecuteChanged();
    }

    private static string DescribeError(Exception exception) => exception switch
    {
        HttpRequestException { StatusCode: HttpStatusCode.Unauthorized } => "HTTP 401 – Authentifizierung fehlgeschlagen.",
        HttpRequestException http when http.StatusCode.HasValue => $"HTTP {(int)http.StatusCode.Value} – {http.StatusCode}.",
        HttpRequestException => "Der KI-Server ist nicht erreichbar.",
        TaskCanceledException => "Zeitüberschreitung beim KI-Request.",
        _ => exception.Message
    };
}
