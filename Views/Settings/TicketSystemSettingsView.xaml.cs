using System;
using System.Windows.Controls;
using System.Windows.Input;
using System.Text.RegularExpressions;

namespace TaskTool.Views.Settings;

public partial class TicketSystemSettingsView : UserControl
{
    private static readonly Regex NonDigitPattern = new("[^0-9]+", RegexOptions.Compiled);
    public TicketSystemSettingsView()
    {
        InitializeComponent();
    }

    private void TicketSystemPasswordTextBox_GotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
        if (sender is TextBox textBox) textBox.SelectAll();
    }

    private void TicketSystemPasswordTextBox_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
    {
        if (sender is not TextBox textBox || textBox.IsKeyboardFocusWithin) return;
        e.Handled = true;
        textBox.Focus();
    }

    private void TicketSystemPasswordTextBox_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
        if (sender is not TextBox textBox) return;
        textBox.Dispatcher.BeginInvoke(new Action(() => textBox.GetBindingExpression(TextBox.TextProperty)?.UpdateTarget()));
    }

    private void NumericTextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        => e.Handled = NonDigitPattern.IsMatch(e.Text);

    private void NumericTextBox_LostFocus(object sender, System.Windows.RoutedEventArgs e)
    {
        if (sender is not TextBox textBox) return;
        var binding = textBox.GetBindingExpression(TextBox.TextProperty);
        binding?.UpdateSource();
        binding?.UpdateTarget();
    }
}
