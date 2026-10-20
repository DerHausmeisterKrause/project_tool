using System.Collections;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using TaskTool.Models;

namespace TaskTool.Views.Controls;

public partial class SearchableTicketFieldPicker : UserControl
{
    private IReadOnlyList<TicketFieldOption> _snapshot = Array.Empty<TicketFieldOption>();

    public static readonly DependencyProperty ItemsSourceProperty = DependencyProperty.Register(
        nameof(ItemsSource), typeof(IEnumerable), typeof(SearchableTicketFieldPicker));

    public static readonly DependencyProperty SelectedItemProperty = DependencyProperty.Register(
        nameof(SelectedItem), typeof(TicketFieldOption), typeof(SearchableTicketFieldPicker),
        new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, OnSelectedItemChanged));

    public static readonly DependencyProperty PlaceholderProperty = DependencyProperty.Register(
        nameof(Placeholder), typeof(string), typeof(SearchableTicketFieldPicker), new PropertyMetadata("Auswählen"));

    public static readonly DependencyProperty SearchPlaceholderProperty = DependencyProperty.Register(
        nameof(SearchPlaceholder), typeof(string), typeof(SearchableTicketFieldPicker), new PropertyMetadata("Suchen …"));

    public IEnumerable? ItemsSource { get => (IEnumerable?)GetValue(ItemsSourceProperty); set => SetValue(ItemsSourceProperty, value); }
    public TicketFieldOption? SelectedItem { get => (TicketFieldOption?)GetValue(SelectedItemProperty); set => SetValue(SelectedItemProperty, value); }
    public string Placeholder { get => (string)GetValue(PlaceholderProperty); set => SetValue(PlaceholderProperty, value); }
    public string SearchPlaceholder { get => (string)GetValue(SearchPlaceholderProperty); set => SetValue(SearchPlaceholderProperty, value); }

    public SearchableTicketFieldPicker()
    {
        InitializeComponent();
        Loaded += (_, _) => UpdateSelectionText();
    }

    internal static bool Matches(TicketFieldOption option, string? searchText)
    {
        var query = searchText?.Trim() ?? string.Empty;
        return query.Length == 0
               || option.Key.Contains(query, StringComparison.OrdinalIgnoreCase)
               || option.DisplayText.Contains(query, StringComparison.OrdinalIgnoreCase);
    }

    internal static IReadOnlyList<TicketFieldOption> Filter(IEnumerable<TicketFieldOption> options, string? searchText)
        => options.Where(option => Matches(option, searchText)).ToArray();

    private void OpenButton_Click(object sender, RoutedEventArgs e)
    {
        _snapshot = ItemsSource?.Cast<object>().OfType<TicketFieldOption>().ToArray()
                    ?? Array.Empty<TicketFieldOption>();
        SearchBox.Text = string.Empty;
        SearchBox.ToolTip = SearchPlaceholder;
        ApplyFilter();
        PickerPopup.IsOpen = true;
        SearchBox.Focus();
    }

    private void SearchBox_TextChanged(object sender, TextChangedEventArgs e) => ApplyFilter();

    private void ApplyFilter()
    {
        if (ResultsList == null) return;
        ResultsList.ItemsSource = Filter(_snapshot, SearchBox?.Text);
        ResultsList.SelectedItem = null;
    }

    private void SearchBox_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Escape) CloseWithoutSelection();
        else if (e.Key == Key.Down && ResultsList.Items.Count > 0)
        {
            ResultsList.SelectedIndex = 0;
            ResultsList.Focus();
            e.Handled = true;
        }
    }

    private void ResultsList_PreviewKeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter && ResultsList.SelectedItem is TicketFieldOption option)
        {
            CommitSelection(option);
            e.Handled = true;
        }
        else if (e.Key == Key.Escape) CloseWithoutSelection();
    }

    private void ResultsList_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
    {
        if (FindAncestor<ListBoxItem>(e.OriginalSource as DependencyObject) is { DataContext: TicketFieldOption option })
            CommitSelection(option);
    }

    internal void CommitSelection(TicketFieldOption option)
    {
        SelectedItem = option;
        PickerPopup.IsOpen = false;
    }

    private void CloseWithoutSelection()
    {
        PickerPopup.IsOpen = false;
    }

    private void PickerPopup_Closed(object? sender, EventArgs e)
    {
        SearchBox.Text = string.Empty;
        ResultsList.ItemsSource = null;
        _snapshot = Array.Empty<TicketFieldOption>();
        UpdateSelectionText();
    }

    private static void OnSelectedItemChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        => ((SearchableTicketFieldPicker)d).UpdateSelectionText();

    private void UpdateSelectionText()
    {
        if (SelectionText != null)
            SelectionText.Text = SelectedItem?.DisplayText ?? Placeholder;
    }

    private static T? FindAncestor<T>(DependencyObject? current) where T : DependencyObject
    {
        while (current != null)
        {
            if (current is T match) return match;
            current = VisualTreeHelper.GetParent(current);
        }
        return null;
    }
}
