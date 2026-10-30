using System.Collections.ObjectModel;
using System.Collections.Specialized;

namespace ExcelMergeFork.App.Ui;

/// <summary>
/// Raises one reset notification after a full replacement, so a large list
/// does not rebuild the visual tree once per row.
/// </summary>
public sealed class RangeObservableCollection<T> : ObservableCollection<T>
{
    private bool _suppress;

    public void Reset(IReadOnlyList<T> items)
    {
        _suppress = true;
        try
        {
            Clear();
            foreach (var item in items)
            {
                Add(item);
            }
        }
        finally
        {
            _suppress = false;
        }

        OnCollectionChanged(new NotifyCollectionChangedEventArgs(NotifyCollectionChangedAction.Reset));
    }

    protected override void OnCollectionChanged(NotifyCollectionChangedEventArgs e)
    {
        if (!_suppress)
        {
            base.OnCollectionChanged(e);
        }
    }
}
