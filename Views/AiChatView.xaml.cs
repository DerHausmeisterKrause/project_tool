using System.Collections.Specialized;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using TaskTool.Models;
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
        Dispatcher.BeginInvoke(() =>
        {
            if (!IsLoaded) return;
            MessageScrollViewer.ScrollToEnd();
            if (args.NewItems?.OfType<AiChatMessage>().Any(message => message.IsUser) == true)
                InputTextBox.Focus();
        }, DispatcherPriority.Loaded);
    }

    private void InputTextBox_OnKeyDown(object sender, KeyEventArgs args)
    {
        if (args.Key != Key.Enter || Keyboard.Modifiers.HasFlag(ModifierKeys.Shift)) return;
        if (DataContext is AiChatViewModel viewModel && viewModel.SendCommand.CanExecute(null)) viewModel.SendCommand.Execute(null);
        args.Handled = true;
    }
}
