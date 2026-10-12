using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Markup;

namespace TaskTool.Infrastructure;

public static class TextInputAssist
{
    public const string SpellcheckContextMenuStyleKey = "SpellcheckContextMenuStyle";
    public const string SpellcheckMenuItemStyleKey = "SpellcheckMenuItemStyle";
    public const string SpellcheckSeparatorStyleKey = "SpellcheckSeparatorStyle";

    public static readonly DependencyProperty IsFreeTextProperty = DependencyProperty.RegisterAttached(
        "IsFreeText", typeof(bool), typeof(TextInputAssist), new PropertyMetadata(false, OnIsFreeTextChanged));
    private static readonly DependencyProperty OriginalContextMenuProperty = DependencyProperty.RegisterAttached(
        "OriginalContextMenu", typeof(ContextMenu), typeof(TextInputAssist));
    private static readonly DependencyProperty SpellcheckContextMenuProperty = DependencyProperty.RegisterAttached(
        "SpellcheckContextMenu", typeof(ContextMenu), typeof(TextInputAssist));

    public static bool GetIsFreeText(DependencyObject element) => (bool)element.GetValue(IsFreeTextProperty);
    public static void SetIsFreeText(DependencyObject element, bool value) => element.SetValue(IsFreeTextProperty, value);

    private static void OnIsFreeTextChanged(DependencyObject element, DependencyPropertyChangedEventArgs args)
    {
        if (element is not TextBoxBase input) return;
        var enabled = args.NewValue is true;
        SpellCheck.SetIsEnabled(input, enabled);
        if (enabled)
        {
            input.Language = XmlLanguage.GetLanguage("de-DE");
            if (input.GetValue(SpellcheckContextMenuProperty) is not ContextMenu)
            {
                input.SetValue(OriginalContextMenuProperty, input.ContextMenu);
                var menu = new ContextMenu();
                menu.SetResourceReference(FrameworkElement.StyleProperty, SpellcheckContextMenuStyleKey);
                input.SetValue(SpellcheckContextMenuProperty, menu);
                input.ContextMenu = menu;
                input.ContextMenuOpening += OnContextMenuOpening;
            }
            return;
        }

        input.ContextMenuOpening -= OnContextMenuOpening;
        if (input.GetValue(SpellcheckContextMenuProperty) is ContextMenu ownedMenu
            && ReferenceEquals(input.ContextMenu, ownedMenu))
            input.ContextMenu = input.GetValue(OriginalContextMenuProperty) as ContextMenu;
        input.ClearValue(SpellcheckContextMenuProperty);
        input.ClearValue(OriginalContextMenuProperty);
    }

    private static void OnContextMenuOpening(object sender, ContextMenuEventArgs args)
    {
        if (sender is TextBoxBase input) RefreshContextMenu(input);
    }

    internal static void RefreshContextMenu(TextBoxBase input)
    {
        if (input.ContextMenu is not ContextMenu menu) return;
        var spellingError = GetSpellingErrorAtPointer(input);
        PopulateContextMenu(input, spellingError?.Suggestions, spellingError != null);
    }

    internal static void PopulateContextMenu(TextBoxBase input, IEnumerable<string>? suggestions, bool hasSpellingError)
    {
        if (input.ContextMenu is not ContextMenu menu) return;
        menu.Items.Clear();
        if (hasSpellingError)
        {
            foreach (var suggestion in suggestions ?? Array.Empty<string>())
                menu.Items.Add(MenuItem(suggestion, EditingCommands.CorrectSpellingError, input, suggestion));
            if (menu.Items.Count > 0) menu.Items.Add(Separator());
            menu.Items.Add(MenuItem("Alles ignorieren", EditingCommands.IgnoreSpellingError, input));
            menu.Items.Add(Separator());
        }

        menu.Items.Add(MenuItem("Ausschneiden", ApplicationCommands.Cut, input, gesture: "Strg+X"));
        menu.Items.Add(MenuItem("Kopieren", ApplicationCommands.Copy, input, gesture: "Strg+C"));
        menu.Items.Add(MenuItem("Einfügen", ApplicationCommands.Paste, input, gesture: "Strg+V"));
    }

    private static SpellingError? GetSpellingErrorAtPointer(TextBoxBase input)
    {
        var point = Mouse.GetPosition(input);
        if (input is TextBox textBox)
        {
            var index = textBox.GetCharacterIndexFromPoint(point, true);
            if (index < 0) return null;
            textBox.CaretIndex = index;
            return textBox.GetSpellingError(index);
        }
        if (input is RichTextBox richTextBox)
        {
            var position = richTextBox.GetPositionFromPoint(point, true);
            if (position == null) return null;
            richTextBox.CaretPosition = position;
            return richTextBox.GetSpellingError(position);
        }
        return null;
    }

    private static MenuItem MenuItem(string header, ICommand command, IInputElement target, object? parameter = null, string gesture = "")
    {
        var item = new MenuItem
        {
            Header = header,
            Command = command,
            CommandTarget = target,
            CommandParameter = parameter,
            InputGestureText = gesture
        };
        item.SetResourceReference(FrameworkElement.StyleProperty, SpellcheckMenuItemStyleKey);
        return item;
    }

    private static Separator Separator()
    {
        var separator = new Separator();
        separator.SetResourceReference(FrameworkElement.StyleProperty, SpellcheckSeparatorStyleKey);
        return separator;
    }
}
