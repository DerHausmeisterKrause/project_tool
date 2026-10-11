using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Markup;

namespace TaskTool.Infrastructure;

public static class TextInputAssist
{
    public static readonly DependencyProperty IsFreeTextProperty = DependencyProperty.RegisterAttached(
        "IsFreeText", typeof(bool), typeof(TextInputAssist), new PropertyMetadata(false, OnIsFreeTextChanged));

    public static bool GetIsFreeText(DependencyObject element) => (bool)element.GetValue(IsFreeTextProperty);
    public static void SetIsFreeText(DependencyObject element, bool value) => element.SetValue(IsFreeTextProperty, value);

    private static void OnIsFreeTextChanged(DependencyObject element, DependencyPropertyChangedEventArgs args)
    {
        if (element is not TextBoxBase input) return;
        var enabled = args.NewValue is true;
        SpellCheck.SetIsEnabled(input, enabled);
        if (enabled) input.Language = XmlLanguage.GetLanguage("de-DE");
    }
}
