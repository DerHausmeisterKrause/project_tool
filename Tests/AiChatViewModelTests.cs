using TaskTool.Models;
using TaskTool.Services;
using TaskTool.ViewModels;
using Xunit;

namespace TaskTool.Tests;

public sealed class AiChatViewModelTests
{
    private static AiChatViewModel CreateViewModel(FakeAi ai)
        => new(ai, () => { });

    [Fact]
    public void InitialState_HasTitleAndCannotSendEmptyText()
    {
        var viewModel = CreateViewModel(new FakeAi());
        Assert.Equal("KI", viewModel.Title);
        Assert.False(viewModel.CanSend);
    }

    [Fact]
    public void DisabledAi_CannotSend()
    {
        var viewModel = CreateViewModel(new FakeAi { IsEnabled = false, CanChat = false });
        viewModel.InputText = "Hallo";
        Assert.False(viewModel.CanSend);
    }

    [Fact]
    public async Task SuccessfulRequest_AddsUserAndAssistantMessages()
    {
        var ai = new FakeAi { Answer = "Guten Tag!" };
        var viewModel = CreateViewModel(ai);
        viewModel.InputText = "Hallo";
        await viewModel.SendAsync();
        Assert.Collection(viewModel.Messages,
            message => Assert.Equal(AiChatRole.User, message.Role),
            message => { Assert.Equal(AiChatRole.Assistant, message.Role); Assert.Equal("Guten Tag!", message.Content); });
        Assert.Equal(0.1, ai.LastOptions!.Temperature);
    }

    [Fact]
    public async Task FailedRequest_DoesNotAddFakeAssistantMessage()
    {
        var viewModel = CreateViewModel(new FakeAi { Exception = new HttpRequestException("kaputt") });
        viewModel.InputText = "Hallo";
        await viewModel.SendAsync();
        Assert.Single(viewModel.Messages);
        Assert.Equal(AiChatRole.User, viewModel.Messages[0].Role);
        Assert.NotEmpty(viewModel.ErrorMessage);
    }

    [Fact]
    public async Task RunningRequest_RejectsSecondRequest()
    {
        var completion = new TaskCompletionSource<string>();
        var ai = new FakeAi { PendingAnswer = completion.Task };
        var viewModel = CreateViewModel(ai);
        viewModel.InputText = "Erste Frage";
        var first = viewModel.SendAsync();
        viewModel.InputText = "Zweite Frage";
        await viewModel.SendAsync();
        Assert.Equal(1, ai.CallCount);
        Assert.Equal(2, viewModel.Messages.Count);
        Assert.True(viewModel.Messages[1].IsTyping);
        Assert.Equal("Plenaro schreibt.", viewModel.Messages[1].Content);
        completion.SetResult("Antwort");
        await first;
    }

    [Fact]
    public async Task RequestContext_IsLimitedToTwentyMessagesPlusSystemPrompt()
    {
        var ai = new FakeAi();
        var viewModel = CreateViewModel(ai);
        for (var index = 0; index < 25; index++)
        {
            viewModel.InputText = $"Frage {index}";
            await viewModel.SendAsync();
        }
        Assert.Equal(AiChatRole.System, ai.LastMessages![0].Role);
        Assert.Equal(AiChatViewModel.MaxContextMessages + 1, ai.LastMessages.Count);
        Assert.DoesNotContain(ai.LastMessages, message => message.Content == "Frage 0");
    }

    [Fact]
    public async Task ClearChat_RemovesInMemoryHistory()
    {
        var viewModel = CreateViewModel(new FakeAi());
        viewModel.InputText = "Hallo";
        await viewModel.SendAsync();
        viewModel.ClearChat();
        Assert.Empty(viewModel.Messages);
    }

    [Fact]
    public void SystemPrompt_RequiresExactOutputWithoutAdditionalText()
    {
        Assert.Contains("exaktes Ausgabeformat", AiChatViewModel.SystemPrompt);
        Assert.Contains("keine Begrüßung, Erklärung, Einleitung oder zusätzlichen Text", AiChatViewModel.SystemPrompt);
    }

    [Fact]
    public void KnowledgeContext_IsMergedIntoSingleSystemMessage()
    {
        var viewModel = CreateViewModel(new FakeAi());
        viewModel.InputText = "Technische Frage";

        var request = viewModel.BuildRequestMessages("LOKALES PLENARO-WISSEN:\nInhalt");

        Assert.Single(request.Where(message => message.Role == AiChatRole.System));
        Assert.Contains("LOKALES PLENARO-WISSEN", request[0].Content);
    }

    private sealed class FakeAi : IAiChatService
    {
        public bool IsEnabled { get; set; } = true;
        public bool CanChat { get; set; } = true;
        public string ProviderDescription => "Test";
        public string AvailabilityMessage => string.Empty;
        public string Answer { get; set; } = "Antwort";
        public Exception? Exception { get; set; }
        public Task<string>? PendingAnswer { get; set; }
        public int CallCount { get; private set; }
        public IReadOnlyList<AiChatRequestMessage>? LastMessages { get; private set; }
        public AiRequestOptions? LastOptions { get; private set; }
        public event EventHandler? StateChanged;

        public Task<string> ChatAsync(IReadOnlyList<AiChatRequestMessage> messages, AiRequestOptions options, CancellationToken cancellationToken = default)
        {
            CallCount++;
            LastMessages = messages;
            LastOptions = options;
            if (Exception != null) return Task.FromException<string>(Exception);
            return PendingAnswer ?? Task.FromResult(Answer);
        }
    }
}
