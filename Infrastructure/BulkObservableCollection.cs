using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;

namespace TaskTool.Infrastructure;

public sealed class BulkObservableCollection<T> : ObservableCollection<T>
{
    public void ReplaceAll(IEnumerable<T> items)
    {
        var replacement = items as IReadOnlyList<T> ?? items.ToArray();
        Items.Clear();
        foreach (var item in replacement)
            Items.Add(item);

        OnPropertyChanged(new PropertyChangedEventArgs(nameof(Count)));
        OnPropertyChanged(new PropertyChangedEventArgs("Item[]"));
        OnCollectionChanged(new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));
    }
}
