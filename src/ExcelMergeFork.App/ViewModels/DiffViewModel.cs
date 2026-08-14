using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ExcelMergeFork.Core.Compare;
using ExcelMergeFork.Core.Settings;

namespace ExcelMergeFork.App.ViewModels;

public sealed partial class DiffViewModel : ObservableObject
{
    private string _local;
    private string _remote;
    private CompareResult? _result;
    private readonly UserSettings _settings = AppSettingsStore.Load();

    public DiffViewModel(string local, string remote)
    {
        _local = local;
        _remote = remote;
        PathLocal = local;
        PathRemote = remote;
        ItemsView = CollectionViewSource.GetDefaultView(Items);
        ItemsView.Filter = FilterItem;
        ShowAdded = _settings.DiffFilter.GetValueOrDefault("新增行", true);
        ShowDeleted = _settings.DiffFilter.GetValueOrDefault("删除行", true);
        ShowNewCols = _settings.DiffFilter.GetValueOrDefault("新增列", true);
        ShowDelCols = _settings.DiffFilter.GetValueOrDefault("删除列", true);
        ShowModified = _settings.DiffFilter.GetValueOrDefault("修改", true);
        AutoOpen = _settings.AutoOpenCompare;
    }

    public ObservableCollection<DiffListItem> Items { get; } = [];
    public ICollectionView ItemsView { get; }

    [ObservableProperty] private string _title = "本地 vs 线上";
    [ObservableProperty] private string _pathLocal = "";
    [ObservableProperty] private string _pathRemote = "";
    [ObservableProperty] private string _statusText = "正在计算差异...";
    [ObservableProperty] private string _searchText = "";
    [ObservableProperty] private bool _busy;
    [ObservableProperty] private bool _showAdded = true;
    [ObservableProperty] private bool _showDeleted = true;
    [ObservableProperty] private bool _showNewCols = true;
    [ObservableProperty] private bool _showDelCols = true;
    [ObservableProperty] private bool _showModified = true;
    [ObservableProperty] private bool _autoOpen;
    [ObservableProperty] private DiffListItem? _selectedItem;

    partial void OnSearchTextChanged(string value) => ItemsView.Refresh();
    partial void OnShowAddedChanged(bool value) { PersistFilter(); ItemsView.Refresh(); }
    partial void OnShowDeletedChanged(bool value) { PersistFilter(); ItemsView.Refresh(); }
    partial void OnShowNewColsChanged(bool value) { PersistFilter(); ItemsView.Refresh(); }
    partial void OnShowDelColsChanged(bool value) { PersistFilter(); ItemsView.Refresh(); }
    partial void OnShowModifiedChanged(bool value) { PersistFilter(); ItemsView.Refresh(); }

    public async Task LoadAsync()
    {
        Busy = true;
        StatusText = "正在计算差异...";
        try
        {
            var local = _local;
            var remote = _remote;
            _result = await Task.Run(() => CompareEngine.Compute(remote, local));
            Items.Clear();
            foreach (var row in _result.Rows)
            {
                Items.Add(new DiffListItem(row));
            }

            StatusText = $"共 {_result.Rows.Count} 行差异，耗时 {_result.ElapsedMs}ms";
            if (AutoOpen && _result.Rows.Count > 0)
            {
                Export(open: true);
            }
        }
        catch (Exception ex)
        {
            StatusText = "对比失败: " + ex.Message;
        }
        finally
        {
            Busy = false;
        }
    }

    [RelayCommand]
    private async Task SwapAsync()
    {
        (_local, _remote) = (_remote, _local);
        PathLocal = _local;
        PathRemote = _remote;
        Title = Title.StartsWith("线上", StringComparison.Ordinal) ? "本地 vs 线上" : "线上 vs 本地";
        await LoadAsync();
    }

    [RelayCommand]
    private void Export(bool open = false)
    {
        if (_result is null)
        {
            return;
        }

        CompareEngine.WriteExcel(_result, open || AutoOpen);
        StatusText = "已导出 " + _result.OutputPath;
    }

    private void PersistFilter()
    {
        _settings.DiffFilter["新增行"] = ShowAdded;
        _settings.DiffFilter["删除行"] = ShowDeleted;
        _settings.DiffFilter["新增列"] = ShowNewCols;
        _settings.DiffFilter["删除列"] = ShowDelCols;
        _settings.DiffFilter["修改"] = ShowModified;
        _settings.AutoOpenCompare = AutoOpen;
        AppSettingsStore.Save(_settings);
    }

    private bool FilterItem(object obj)
    {
        if (obj is not DiffListItem row)
        {
            return false;
        }

        var visible = row.Status switch
        {
            "新增行" => ShowAdded,
            "删除行" => ShowDeleted,
            "新增列" => ShowNewCols,
            "删除列" => ShowDelCols,
            "修改" => ShowModified,
            _ => true,
        };
        if (!visible)
        {
            return false;
        }

        if (string.IsNullOrWhiteSpace(SearchText))
        {
            return true;
        }

        return row.Key.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
               row.Sheet.Contains(SearchText, StringComparison.OrdinalIgnoreCase);
    }
}

public sealed class DiffListItem
{
    public DiffListItem(DiffRow row)
    {
        Sheet = row.Sheet;
        Key = row.Key;
        Status = row.Status;
        Left = row.Left;
        Right = row.Right;
        TypeLabel = row.Status.Replace("行", "").Replace("列", "");
    }

    public string Sheet { get; }
    public string Key { get; }
    public string Status { get; }
    public string Left { get; }
    public string Right { get; }
    public string TypeLabel { get; }
    public IReadOnlyList<string> LeftValues => Left.Split(" | ");
    public IReadOnlyList<string> RightValues => Right.Split(" | ");
}
