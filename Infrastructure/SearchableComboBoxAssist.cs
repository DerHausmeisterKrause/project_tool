using System.Collections.Specialized;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using TaskTool.Models;

namespace TaskTool.Infrastructure;

public static class SearchableComboBoxAssist
{
    public static readonly DependencyProperty IsEnabledProperty = DependencyProperty.RegisterAttached(
        "IsEnabled", typeof(bool), typeof(SearchableComboBoxAssist), new PropertyMetadata(false, OnIsEnabledChanged));
    private static readonly DependencyProperty StateProperty = DependencyProperty.RegisterAttached(
        "State", typeof(State), typeof(SearchableComboBoxAssist));

    public static bool GetIsEnabled(DependencyObject element) => (bool)element.GetValue(IsEnabledProperty);
    public static void SetIsEnabled(DependencyObject element, bool value) => element.SetValue(IsEnabledProperty, value);

    internal static bool Matches(TicketFieldOption option, string? searchText)
    {
        var query = searchText?.Trim() ?? string.Empty;
        return query.Length == 0
               || option.Key.Contains(query, StringComparison.OrdinalIgnoreCase)
               || option.DisplayText.Contains(query, StringComparison.OrdinalIgnoreCase);
    }

    private static void OnIsEnabledChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is not ComboBox comboBox) return;
        if ((bool)e.NewValue)
        {
            var state = new State(comboBox);
            comboBox.SetValue(StateProperty, state);
            state.Attach();
        }
        else if (comboBox.GetValue(StateProperty) is State state)
        {
            state.Detach();
            comboBox.ClearValue(StateProperty);
        }
    }

    private sealed class State
    {
        private readonly ComboBox _comboBox;
        private TextBox? _editor;
        private object? _selectionBeforeSearch;
        private bool _selectionCommitted;
        private bool _suppressText;
        private INotifyCollectionChanged? _observableItems;

        public State(ComboBox comboBox) => _comboBox = comboBox;

        public void Attach()
        {
            _comboBox.IsEditable = true;
            _comboBox.IsTextSearchEnabled = false;
            _comboBox.StaysOpenOnEdit = true;
            _comboBox.Loaded += OnLoaded;
            _comboBox.Unloaded += OnUnloaded;
            _comboBox.DropDownOpened += OnDropDownOpened;
            _comboBox.DropDownClosed += OnDropDownClosed;
            _comboBox.SelectionChanged += OnSelectionChanged;
            _comboBox.PreviewKeyDown += OnPreviewKeyDown;
            SubscribeItems();
        }

        public void Detach()
        {
            DetachEditor();
            _comboBox.Loaded -= OnLoaded;
            _comboBox.Unloaded -= OnUnloaded;
            _comboBox.DropDownOpened -= OnDropDownOpened;
            _comboBox.DropDownClosed -= OnDropDownClosed;
            _comboBox.SelectionChanged -= OnSelectionChanged;
            _comboBox.PreviewKeyDown -= OnPreviewKeyDown;
            if (_observableItems != null) _observableItems.CollectionChanged -= OnItemsChanged;
            ResetFilter();
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            _comboBox.ApplyTemplate();
            _editor = _comboBox.Template.FindName("PART_EditableTextBox", _comboBox) as TextBox;
            if (_editor != null) _editor.TextChanged += OnTextChanged;
            SubscribeItems();
            ShowSelectionText();
        }

        private void OnUnloaded(object sender, RoutedEventArgs e) => DetachEditor();

        private void OnDropDownOpened(object? sender, EventArgs e)
        {
            _selectionBeforeSearch = _comboBox.SelectedItem;
            _selectionCommitted = false;
            SetEditorText(string.Empty);
            ResetFilter();
            _editor?.Focus();
        }

        private void OnDropDownClosed(object? sender, EventArgs e)
        {
            ResetFilter();
            if (!_selectionCommitted && _selectionBeforeSearch != null)
                _comboBox.SelectedItem = _selectionBeforeSearch;
            ShowSelectionText();
        }

        private void OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_comboBox.IsDropDownOpen && e.AddedItems.OfType<TicketFieldOption>().Any())
            {
                _selectionCommitted = true;
                _selectionBeforeSearch = _comboBox.SelectedItem;
            }
        }

        private void OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.Escape) return;
            ResetFilter();
            SetEditorText(string.Empty);
            _comboBox.IsDropDownOpen = false;
            e.Handled = true;
        }

        private void OnTextChanged(object sender, TextChangedEventArgs e)
        {
            if (_suppressText || !_comboBox.IsDropDownOpen) return;
            var query = _editor?.Text ?? string.Empty;
            if (_comboBox.ItemsSource == null) return;
            var view = CollectionViewSource.GetDefaultView(_comboBox.ItemsSource);
            if (view == null) return;
            view.Filter = item => item is TicketFieldOption option && Matches(option, query);
            view.Refresh();
        }

        private void OnItemsChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            ResetFilter();
            if (!_comboBox.IsDropDownOpen) ShowSelectionText();
        }

        private void SubscribeItems()
        {
            if (ReferenceEquals(_observableItems, _comboBox.ItemsSource)) return;
            if (_observableItems != null) _observableItems.CollectionChanged -= OnItemsChanged;
            _observableItems = _comboBox.ItemsSource as INotifyCollectionChanged;
            if (_observableItems != null) _observableItems.CollectionChanged += OnItemsChanged;
        }

        private void ResetFilter()
        {
            if (_comboBox.ItemsSource == null) return;
            var view = CollectionViewSource.GetDefaultView(_comboBox.ItemsSource);
            if (view != null) view.Filter = null;
        }

        private void ShowSelectionText()
            => SetEditorText((_comboBox.SelectedItem as TicketFieldOption)?.DisplayText ?? string.Empty);

        private void SetEditorText(string text)
        {
            if (_editor == null) return;
            _suppressText = true;
            _editor.Text = text;
            _editor.CaretIndex = text.Length;
            _suppressText = false;
        }

        private void DetachEditor()
        {
            if (_editor != null) _editor.TextChanged -= OnTextChanged;
            _editor = null;
        }
    }
}
