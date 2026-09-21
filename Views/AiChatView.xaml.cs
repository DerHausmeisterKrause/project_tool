using System.Collections.Specialized;
using System.Windows.Controls;
using System.Windows.Input;
using TaskTool.ViewModels;

namespace TaskTool.Views;

public partial class AiChatView : UserControl
{
    public AiChatView()
    {
        InitializeComponent();
        DataContextChanged += (_, args) =>
        {
            if (args.OldValue is AiChatViewModel oldViewModel) oldViewModel.Messages.CollectionChanged -= MessagesOnCollectionChanged;
            if (args.NewValue is AiChatViewModel newViewModel) newViewModel.Messages.CollectionChanged += MessagesOnCollectionChanged;
        };
    }

    private void MessagesOnCollectionChanged(object? sender, NotifyCollectionChangedEventArgs args)
    {
        if (MessageList.Items.Count > 0) MessageList.ScrollIntoView(MessageList.Items[^1]);
    }

    private void InputTextBox_OnKeyDown(object sender, KeyEventArgs args)
    {
        if (args.Key != Key.Enter || Keyboard.Modifiers.HasFlag(ModifierKeys.Shift)) return;
        if (DataContext is AiChatViewModel viewModel && viewModel.SendCommand.CanExecute(null)) viewModel.SendCommand.Execute(null);
        args.Handled = true;
    }
}
