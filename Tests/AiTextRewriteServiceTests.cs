using System.Windows.Controls;
using System.Windows;
using System.Windows.Markup;
using System.Windows.Media;
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
        Assert.Equal("de-DE", language!.IetfLanguageTag, StringComparer.OrdinalIgnoreCase);
    }

    [Fact]
    public void FreeTextBehavior_PreservesExistingContextMenuAndDisablesSpellcheck()
    {
        Exception? error = null;
        var thread = new Thread(() =>
        {
            try
            {
                var original = new ContextMenu();
                var textBox = new TextBox { ContextMenu = original };
                TextInputAssist.SetIsFreeText(textBox, true);
                Assert.Same(original, textBox.ContextMenu);

                TextInputAssist.SetIsFreeText(textBox, false);
                Assert.False(SpellCheck.GetIsEnabled(textBox));
                Assert.Same(original, textBox.ContextMenu);
            }
            catch (Exception exception) { error = exception; }
        });
        thread.SetApartmentState(ApartmentState.STA); thread.Start(); thread.Join();
        Assert.Null(error);
    }

    [Fact]
    public void FreeTextBehavior_DoesNotInstallAContextMenu()
    {
        Exception? error = null;
        var thread = new Thread(() =>
        {
            try
            {
                var textBox = new TextBox();
                TextInputAssist.SetIsFreeText(textBox, true);
                Assert.Null(textBox.ContextMenu);
            }
            catch (Exception exception) { error = exception; }
        });
        thread.SetApartmentState(ApartmentState.STA); thread.Start(); thread.Join();
        Assert.Null(error);
    }

    [Fact]
    public void Theme_LetsSuggestionTextBlocksInheritNativeMenuColors()
    {
        Exception? error = null;
        var thread = new Thread(() =>
        {
            try
            {
                var resources = new ResourceDictionary
                {
                    Source = new Uri("/TaskTool;component/Themes/Theme.xaml", UriKind.Relative)
                };
                Assert.False(resources.Contains(typeof(TextBlock)));
                var userControlStyle = Assert.IsType<Style>(resources[typeof(UserControl)]);
                Assert.Contains(userControlStyle.Setters.OfType<Setter>(),
                    setter => setter.Property == Control.ForegroundProperty);
                Assert.Equal(Color.FromRgb(0x37, 0x41, 0x51),
                    Assert.IsType<SolidColorBrush>(resources[SystemColors.MenuTextBrushKey]).Color);
                Assert.Equal(Color.FromRgb(0x6B, 0x72, 0x80),
                    Assert.IsType<SolidColorBrush>(resources[SystemColors.GrayTextBrushKey]).Color);
            }
            catch (Exception exception) { error = exception; }
        });
        thread.SetApartmentState(ApartmentState.STA); thread.Start(); thread.Join();
        Assert.Null(error);
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
