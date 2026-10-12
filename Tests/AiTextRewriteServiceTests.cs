using System.Windows.Controls;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;
using TaskTool.Infrastructure;
using TaskTool.Models;
using TaskTool.Services;
using Xunit;

namespace TaskTool.Tests;

public sealed class AiTextRewriteServiceTests
{
    [Fact]
    public async Task Rewrite_UsesStrictPromptAndLowTemperature()
    {
        var ai = new FakeAi { Answer = "Hallo Herr Müller, das Problem ist behoben." };
        var service = new AiTextRewriteService(ai, new LoggerService());

        var result = await service.RewriteCustomerReplyAsync("hallo herr müller das problem ist behoben");

        Assert.Equal(ai.Answer, result);
        Assert.Equal(0.15, ai.Options!.Temperature);
        Assert.Contains("keine neuen fachlichen Informationen", ai.Messages![0].Content, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("ausschließlich mit dem überarbeiteten Text", ai.Messages[0].Content, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("PLENARO-WISSEN", ai.Messages[0].Content, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Rewrite_ProtectsTemplateVariablesUrlsAndEmailAddresses()
    {
        var ai = new FakeAi { Answer = "Hallo __PLENARO_PROTECTED_0__,\n\nbitte teste __PLENARO_PROTECTED_1__\nKontakt: __PLENARO_PROTECTED_2__\nTicket: __PLENARO_PROTECTED_3__" };
        var service = new AiTextRewriteService(ai, new LoggerService());
        const string draft = "Hallo $CustomerFirstName, bitte teste https://srv01.example.local/portal Kontakt: admin@example.local Ticket: $TicketNumber";

        var result = await service.RewriteCustomerReplyAsync(draft);

        Assert.Contains("$CustomerFirstName", result);
        Assert.Contains("https://srv01.example.local/portal", result);
        Assert.Contains("admin@example.local", result);
        Assert.Contains("$TicketNumber", result);
        Assert.DoesNotContain("$CustomerFirstName", ai.Messages![1].Content);
        Assert.DoesNotContain("https://srv01.example.local/portal", ai.Messages[1].Content);
    }

    [Fact]
    public async Task Rewrite_RejectsEmptyResponseOrMissingProtectedMarker()
    {
        var service = new AiTextRewriteService(new FakeAi { Answer = "   " }, new LoggerService());
        await Assert.ThrowsAsync<InvalidDataException>(() => service.RewriteCustomerReplyAsync("Text"));

        service = new AiTextRewriteService(new FakeAi { Answer = "Der Platzhalter wurde entfernt." }, new LoggerService());
        await Assert.ThrowsAsync<InvalidDataException>(() => service.RewriteCustomerReplyAsync("Hallo $Customer"));
    }

    [Fact]
    public async Task Rewrite_PropagatesProviderFailureWithoutReturningReplacementText()
    {
        var service = new AiTextRewriteService(new FakeAi { Error = new HttpRequestException("offline") }, new LoggerService());
        await Assert.ThrowsAsync<HttpRequestException>(() => service.RewriteCustomerReplyAsync("Originaltext"));
    }

    [Fact]
    public void FreeTextBehavior_EnablesGermanWpfSpellCheck()
    {
        Exception? error = null; bool enabled = false; XmlLanguage? language = null;
        var thread = new Thread(() =>
        {
            try
            {
                var textBox = new TextBox();
                TextInputAssist.SetIsFreeText(textBox, true);
                enabled = SpellCheck.GetIsEnabled(textBox); language = textBox.Language;
            }
            catch (Exception exception) { error = exception; }
        });
        thread.SetApartmentState(ApartmentState.STA); thread.Start(); thread.Join();

        Assert.Null(error);
        Assert.True(enabled);
        Assert.NotNull(language);
        Assert.True(string.Equals("de-DE", language.IetfLanguageTag, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void FreeTextBehavior_ProvidesStyledNativeCommandMenuAndRestoresOriginalMenu()
    {
        Exception? error = null;
        var thread = new Thread(() =>
        {
            try
            {
                var original = new ContextMenu();
                var textBox = new TextBox { ContextMenu = original, Text = "korrekt" };
                TextInputAssist.SetIsFreeText(textBox, true);

                var menu = Assert.IsType<ContextMenu>(textBox.ContextMenu);
                Assert.NotSame(original, menu);
                Assert.NotEqual(DependencyProperty.UnsetValue, menu.ReadLocalValue(FrameworkElement.StyleProperty));
                TextInputAssist.RefreshContextMenu(textBox);
                var commandItems = menu.Items.OfType<MenuItem>().ToList();
                Assert.Contains(commandItems, item => item.Command == ApplicationCommands.Cut);
                Assert.Contains(commandItems, item => item.Command == ApplicationCommands.Copy);
                Assert.Contains(commandItems, item => item.Command == ApplicationCommands.Paste);
                Assert.All(commandItems, item => Assert.NotEqual(DependencyProperty.UnsetValue, item.ReadLocalValue(FrameworkElement.StyleProperty)));
                Assert.Empty(menu.Items.OfType<Separator>());

                TextInputAssist.SetIsFreeText(textBox, false);
                Assert.False(SpellCheck.GetIsEnabled(textBox));
                Assert.Same(original, textBox.ContextMenu);
                Assert.Empty(menu.Items);
                menu.Items.Add(new MenuItem { Header = "lifetime probe" });
                menu.RaiseEvent(new RoutedEventArgs(ContextMenu.ClosedEvent));
                Assert.Single(menu.Items);
                menu.Items.Clear();
            }
            catch (Exception exception) { error = exception; }
        });
        thread.SetApartmentState(ApartmentState.STA); thread.Start(); thread.Join();
        Assert.Null(error);
    }

    [Fact]
    public void SpellcheckMenu_UsesNativeCorrectionAndIgnoreCommandsForSuggestions()
    {
        Exception? error = null;
        var thread = new Thread(() =>
        {
            try
            {
                var textBox = new TextBox();
                TextInputAssist.SetIsFreeText(textBox, true);
                TextInputAssist.PopulateContextMenu(textBox, ["Test", "Tat"], hasSpellingError: true);

                var menu = Assert.IsType<ContextMenu>(textBox.ContextMenu);
                var items = menu.Items.OfType<MenuItem>().ToList();
                var suggestions = items.Where(item => item.Command == EditingCommands.CorrectSpellingError).ToList();
                Assert.Equal(["Test", "Tat"], suggestions.Select(item => item.CommandParameter));
                Assert.All(suggestions, item => Assert.Same(textBox, item.CommandTarget));
                Assert.Contains(items, item => item.Command == EditingCommands.IgnoreSpellingError && ReferenceEquals(item.CommandTarget, textBox));
                Assert.Equal(2, menu.Items.OfType<Separator>().Count());
            }
            catch (Exception exception) { error = exception; }
        });
        thread.SetApartmentState(ApartmentState.STA); thread.Start(); thread.Join();
        Assert.Null(error);
    }

    [Fact]
    public void SpellcheckMenu_LimitsSuggestionsAndReleasesItemsWhenClosed()
    {
        Exception? error = null;
        var thread = new Thread(() =>
        {
            try
            {
                var textBox = new TextBox();
                TextInputAssist.SetIsFreeText(textBox, true);
                var menu = Assert.IsType<ContextMenu>(textBox.ContextMenu);
                var suggestions = Enumerable.Range(1, 100).Select(index => $"Vorschlag {index}").ToArray();

                TextInputAssist.PopulateContextMenu(textBox, suggestions, hasSpellingError: true);
                Assert.Equal(TextInputAssist.MaxSpellingSuggestions,
                    menu.Items.OfType<MenuItem>().Count(item => item.Command == EditingCommands.CorrectSpellingError));

                menu.RaiseEvent(new RoutedEventArgs(ContextMenu.ClosedEvent));
                Assert.Empty(menu.Items);

                TextInputAssist.PopulateContextMenu(textBox, ["Neu"], hasSpellingError: true);
                Assert.Contains(menu.Items.OfType<MenuItem>(),
                    item => item.Command == EditingCommands.CorrectSpellingError && Equals(item.CommandParameter, "Neu"));
            }
            catch (Exception exception) { error = exception; }
        });
        thread.SetApartmentState(ApartmentState.STA); thread.Start(); thread.Join();
        Assert.Null(error);
    }

    [Fact]
    public void SpellcheckSuggestions_AreEnumeratedOnlyUpToTheLimit()
    {
        var source = new CountingSuggestions(1000);

        var result = TextInputAssist.MaterializeSuggestions(source);

        Assert.Equal(TextInputAssist.MaxSpellingSuggestions, result.Count);
        Assert.Equal(TextInputAssist.MaxSpellingSuggestions, source.YieldCount);
    }

    private sealed class CountingSuggestions(int count) : IEnumerable<string>
    {
        public int YieldCount { get; private set; }

        public IEnumerator<string> GetEnumerator()
        {
            for (var index = 0; index < count; index++)
            {
                YieldCount++;
                yield return $"Vorschlag {index}";
            }
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
    }

    private sealed class FakeAi : IAiChatService
    {
        public bool IsEnabled { get; set; } = true;
        public bool CanChat { get; set; } = true;
        public string ProviderDescription => "Fake";
        public string AvailabilityMessage => "";
        public string Answer { get; set; } = "";
        public Exception? Error { get; set; }
        public IReadOnlyList<AiChatRequestMessage>? Messages { get; private set; }
        public AiRequestOptions? Options { get; private set; }
        public event EventHandler? StateChanged { add { } remove { } }
        public Task<string> ChatAsync(IReadOnlyList<AiChatRequestMessage> messages, AiRequestOptions options, CancellationToken cancellationToken = default)
        {
            Messages = messages; Options = options;
            return Error == null ? Task.FromResult(Answer) : Task.FromException<string>(Error);
        }
    }
}
