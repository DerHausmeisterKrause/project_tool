using System;
using System.Windows.Controls;
using System.Windows.Input;

namespace TaskTool.Views;

public partial class SettingsView : UserControl
{
    public SettingsView()
    {
        InitializeComponent();
    }

    private void TicketSystemPasswordTextBox_GotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
    {
        if (sender is TextBox textBox)
            textBox.SelectAll();
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

        textBox.Dispatcher.BeginInvoke(new Action(() =>
            textBox.GetBindingExpression(TextBox.TextProperty)?.UpdateTarget()));
    }
}
