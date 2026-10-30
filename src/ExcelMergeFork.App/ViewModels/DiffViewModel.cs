using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;
using System.Windows.Media;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ExcelMergeFork.App.Ui;
using ExcelMergeFork.Core.Compare;
using ExcelMergeFork.Core.Settings;

namespace ExcelMergeFork.App.ViewModels;

public sealed partial class DiffViewModel : ObservableObject
{
    private static readonly SolidColorBrush NewBg = Freeze(Color.FromRgb(0xDC, 0xFC, 0xE7));
    private static readonly SolidColorBrush NewFg = Freeze(Color.FromRgb(0x15, 0x80, 0x3D));
    private static readonly SolidColorBrush DeleteBg = Freeze(Color.FromRgb(0xFE, 0xE2, 0xE2));
    private static readonly SolidColorBrush DeleteFg = Freeze(Color.FromRgb(0xB9, 0x1C, 0x1C));
    private static readonly SolidColorBrush ModifyBg = Freeze(Color.FromRgb(0xFE, 0xF3, 0xC7));
    private static readonly SolidColorBrush ModifyFg = Freeze(Color.FromRgb(0xB4, 0x53, 0x09));

    private string _local;
    private string _remote;
    private CompareResult? _result;
    private int _elapsedMs;
    private bool _suppress;
    private bool _ready;
    private readonly UiDelay _searchDelay = new();
    private readonly UserSettings _settings = AppSettingsStore.Load();

    public DiffViewModel(string local, string remote)
    {
        _local = local;
        _remote = remote;
        PathLocal = local;
        PathRemote = remote;
        FileLocal = System.IO.Path.GetFileName(local);
        FileRemote = System.IO.Path.GetFileName(remote);
        ItemsView = CollectionViewSource.GetDefaultView(Items);
        ItemsView.Filter = FilterItem;
        ShowAdded = _settings.DiffFilter.GetValueOrDefault("新增行", true);
        ShowDeleted = _settings.DiffFilter.GetValueOrDefault("删除行", true);
        ShowNewCols = _settings.DiffFilter.GetValueOrDefault("新增列", true);
        ShowDelCols = _settings.DiffFilter.GetValueOrDefault("删除列", true);
        ShowModified = _settings.DiffFilter.GetValueOrDefault("修改", true);
        AutoOpen = _settings.AutoOpenCompare;
        _ready = true;
        EmptyTitle = "正在对比";
        EmptyMessage = "正在按工作表读取两边的单元格。";
        ShowEmpty = true;
    }

    public RangeObservableCollection<DiffListItem> Items { get; } = [];
    public ObservableCollection<SheetOption> Sheets { get; } = [];
    public ICollectionView ItemsView { get; }

    [ObservableProperty] private string _pathLocal = "";
    [ObservableProperty] private string _pathRemote = "";
    [ObservableProperty] private string _fileLocal = "";
    [ObservableProperty] private string _fileRemote = "";
    [ObservableProperty] private string _statusText = "正在计算差异...";
    [ObservableProperty] private string _searchText = "";
    [ObservableProperty] private bool _busy;
    [ObservableProperty] private bool _showAdded = true;
    [ObservableProperty] private bool _showDeleted = true;
    [ObservableProperty] private bool _showNewCols = true;
    [ObservableProperty] private bool _showDelCols = true;
    [ObservableProperty] private bool _showModified = true;
    [ObservableProperty] private bool _autoOpen;
    [ObservableProperty] private bool _showAllCells;
    [ObservableProperty] private bool _showEmpty = true;
    [ObservableProperty] private bool _hasSelection;
    [ObservableProperty] private bool _showSkipped;
    [ObservableProperty] private string _emptyTitle = "";
    [ObservableProperty] private string _emptyMessage = "";
    [ObservableProperty] private string _skippedHint = "";
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(DetailCells))]
    private DiffListItem? _selectedItem;
    [ObservableProperty] private SheetOption? _selectedSheet;

    public IReadOnlyList<DiffCell> DetailCells =>
        SelectedItem is null
            ? []
            : ShowAllCells
                ? SelectedItem.Cells
                : SelectedItem.Cells.Where(cell => cell.Changed).ToList();

    partial void OnSelectedItemChanged(DiffListItem? value) => HasSelection = value is not null;
    partial void OnAutoOpenChanged(bool value) => PersistFilter();
    partial void OnShowAllCellsChanged(bool value) => OnPropertyChanged(nameof(DetailCells));
    partial void OnSearchTextChanged(string value) => _searchDelay.Schedule(RefreshItems);

    public void Release() => _searchDelay.Dispose();
    partial void OnSelectedSheetChanged(SheetOption? value) => RefreshItems();
    partial void OnShowAddedChanged(bool value) { PersistFilter(); RefreshItems(); }
    partial void OnShowDeletedChanged(bool value) { PersistFilter(); RefreshItems(); }
    partial void OnShowNewColsChanged(bool value) { PersistFilter(); RefreshItems(); }
    partial void OnShowDelColsChanged(bool value) { PersistFilter(); RefreshItems(); }
    partial void OnShowModifiedChanged(bool value) { PersistFilter(); RefreshItems(); }

    public async Task LoadAsync()
    {
        Busy = true;
        StatusText = "正在计算差异...";
        ShowEmpty = true;
        EmptyTitle = "正在对比";
        EmptyMessage = "正在按工作表读取两边的单元格。";
        var autoExported = false;
        try
        {
            var local = _local;
            var remote = _remote;
            // 左边是线上，右边是本地。界面上的「本地 / 线上」按这个方向显示。
            _result = await Task.Run(() => CompareEngine.Compute(remote, local));
            _elapsedMs = _result.ElapsedMs;
            var rows = _result.Rows;
            var built = await Task.Run(() =>
            {
                var list = new List<DiffListItem>(rows.Count);
                foreach (var row in rows)
                {
                    list.Add(new DiffListItem(row, NewBg, NewFg, DeleteBg, DeleteFg, ModifyBg, ModifyFg));
                }

                return list;
            });
            SelectedItem = null;
            Items.Reset(built);

            BuildSheets();
            ShowSkipped = _result.SkippedSheetNames.Count > 0;
            SkippedHint = ShowSkipped
                ? "已跳过以 # 开头的工作表：" + string.Join("、", _result.SkippedSheetNames)
                : "";
            autoExported = AutoOpen && _result.Rows.Count > 0;
            if (autoExported)
            {
                Export(open: true);
            }
        }
        catch (Exception ex)
        {
            StatusText = "对比失败: " + ex.Message;
            ShowEmpty = true;
            EmptyTitle = "对比失败";
            EmptyMessage = ex.Message;
            autoExported = true;
        }
        finally
        {
            Busy = false;
            if (_result is not null && !autoExported)
            {
                UpdateStatus();
            }
        }
    }

    [RelayCommand]
    private async Task SwapAsync()
    {
        (_local, _remote) = (_remote, _local);
        PathLocal = _local;
        PathRemote = _remote;
        FileLocal = System.IO.Path.GetFileName(_local);
        FileRemote = System.IO.Path.GetFileName(_remote);
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

    private void BuildSheets()
    {
        if (_result is null)
        {
            return;
        }

        var counts = _result.Rows
            .GroupBy(row => row.Sheet, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        var options = new List<SheetOption>
        {
            new() { Name = "全部工作表", Sheet = null, Count = _result.Rows.Count },
        };
        foreach (var name in _result.SheetNames)
        {
            counts.TryGetValue(name, out var count);
            options.Add(new SheetOption { Name = name, Sheet = name, Count = count });
        }

        _suppress = true;
        Sheets.Clear();
        foreach (var option in options)
        {
            Sheets.Add(option);
        }

        SelectedSheet = options.FirstOrDefault(option => option.Sheet is not null && option.Count > 0) ?? options[0];
        _suppress = false;
        RefreshItems();
    }

    private void RefreshItems()
    {
        if (_suppress || ItemsView is null)
        {
            return;
        }

        ItemsView.Refresh();
        if (SelectedItem is not null && !FilterItem(SelectedItem))
        {
            SelectedItem = null;
        }

        UpdateStatus();
    }

    private void UpdateStatus()
    {
        if (_result is null)
        {
            return;
        }

        var sheet = SelectedSheet?.Sheet;
        var scoped = sheet is null
            ? _result.Rows.Count
            : _result.Rows.Count(row => row.Sheet == sheet);
        var shown = ItemsView is CollectionView collection ? collection.Count : ItemsView.Cast<object>().Count();

        var scopeName = sheet ?? "全部工作表";
        StatusText = $"{scopeName}：{scoped} 处差异，当前显示 {shown} 处。全部 {_result.Rows.Count} 处，耗时 {_elapsedMs} ms";
        if (_result.Rows.Count == 0)
        {
            ShowEmpty = true;
            EmptyTitle = "没有差异";
            EmptyMessage = _result.SkippedSheetNames.Count > 0 && _result.SheetNames.Count == 0
                ? "这些工作表以 # 开头，对比时会跳过：" + string.Join("、", _result.SkippedSheetNames)
                : "按第一列主键逐条对比后，两边内容一致。";
            return;
        }

        if (shown == 0)
        {
            ShowEmpty = true;
            if (scoped == 0)
            {
                EmptyTitle = "这个工作表没有差异";
                EmptyMessage = "左侧可以切换到其他工作表。有差异的表会标出数量。";
            }
            else
            {
                EmptyTitle = "没有符合筛选的差异";
                EmptyMessage = "试试清空搜索，或勾选上方的差异类型。";
            }

            return;
        }

        ShowEmpty = false;
    }

    private void PersistFilter()
    {
        if (!_ready)
        {
            return;
        }

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

        if (SelectedSheet?.Sheet is { Length: > 0 } sheet && !string.Equals(row.Sheet, sheet, StringComparison.Ordinal))
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
               row.Sheet.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
               row.Left.Contains(SearchText, StringComparison.OrdinalIgnoreCase) ||
               row.Right.Contains(SearchText, StringComparison.OrdinalIgnoreCase);
    }

    private static SolidColorBrush Freeze(Color color)
    {
        var brush = new SolidColorBrush(color);
        brush.Freeze();
        return brush;
    }
}

public sealed class SheetOption
{
    public required string Name { get; init; }
    public string? Sheet { get; init; }
    public int Count { get; init; }
    public string Badge => Count == 0 ? "无差异" : Count + " 处";
}

public sealed class DiffListItem
{
    public DiffListItem(
        DiffRow row,
        SolidColorBrush newBg,
        SolidColorBrush newFg,
        SolidColorBrush deleteBg,
        SolidColorBrush deleteFg,
        SolidColorBrush modifyBg,
        SolidColorBrush modifyFg)
    {
        Sheet = row.Sheet;
        Key = row.Key;
        Status = row.Status;
        Left = row.Left;
        Right = row.Right;
        Cells = row.Cells;
        Location = FormatLocation(row.LeftRow, row.RightRow);
        var changed = row.Cells.Where(cell => cell.Changed).ToList();
        Summary = Summarize(changed);
        Preview = PreviewOf(changed);
        (BadgeBackground, BadgeForeground) = row.Status switch
        {
            "新增行" or "新增列" => (newBg, newFg),
            "删除行" or "删除列" => (deleteBg, deleteFg),
            _ => (modifyBg, modifyFg),
        };
    }

    public string Sheet { get; }
    public string Key { get; }
    public string Status { get; }
    public string Left { get; }
    public string Right { get; }
    public string Location { get; }
    public string Summary { get; }
    public string Preview { get; }
    public IReadOnlyList<DiffCell> Cells { get; }
    public SolidColorBrush BadgeBackground { get; }
    public SolidColorBrush BadgeForeground { get; }

    private static string FormatLocation(int remoteRow, int localRow)
    {
        if (localRow > 0 && remoteRow > 0 && localRow != remoteRow)
        {
            return $"本地 {localRow} 行 · 线上 {remoteRow} 行";
        }

        if (localRow > 0 && remoteRow > 0)
        {
            return $"第 {localRow} 行";
        }

        if (localRow > 0)
        {
            return $"本地第 {localRow} 行";
        }

        if (remoteRow > 0)
        {
            return $"线上第 {remoteRow} 行";
        }

        return "";
    }

    private static string Summarize(IReadOnlyList<DiffCell> changed)
    {
        if (changed.Count == 0)
        {
            return "";
        }

        var names = changed.Select(cell => cell.Header).Where(name => name.Length > 0).Take(3);
        var text = string.Join("、", names);
        if (changed.Count > 3)
        {
            text += " 等 " + changed.Count + " 列";
        }

        return text;
    }

    private static string PreviewOf(IReadOnlyList<DiffCell> changed)
    {
        if (changed.Count == 0)
        {
            return "";
        }

        var cell = changed[0];
        if (changed.Count > 1 && cell.Left.Length == 0 && cell.Right.Length > 0 && cell.Header is "引导ID" or "id" or "Key")
        {
            cell = changed[1];
        }

        var left = Short(cell.Left);
        var right = Short(cell.Right);
        if (left.Length == 0)
        {
            return cell.Header + "： " + right;
        }

        if (right.Length == 0)
        {
            return cell.Header + "： " + left;
        }

        return cell.Header + "： " + left + "  →  " + right;
    }

    private static string Short(string value)
    {
        var text = value.Replace("\r", " ").Replace("\n", " ").Trim();
        return text.Length <= 42 ? text : text[..42] + "…";
    }
}
