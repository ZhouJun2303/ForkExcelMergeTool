using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows.Data;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using ExcelMergeFork.Core.Backup;
using ExcelMergeFork.Core.Excel;
using ExcelMergeFork.Core.Git;
using ExcelMergeFork.Core.Merge;
using ExcelMergeFork.Core.Settings;

namespace ExcelMergeFork.App.ViewModels;

public sealed partial class MergeViewModel : ObservableObject
{
    private readonly string _local;
    private readonly string _base;
    private readonly string _remote;
    private readonly string _merged;
    private readonly GitDriverRequest? _driver;
    private readonly MergeWizardState _wizard = new();
    private MergeSession? _session;
    private MergePreview? _preview;
    private List<PreviewRow> _allItems = [];
    private UserSettings _settings = AppSettingsStore.Load();

    public MergeViewModel(string local, string basePath, string remote, string merged, GitDriverRequest? driver = null)
    {
        _local = local;
        _base = basePath;
        _remote = remote;
        _merged = merged;
        _driver = driver;
        ItemsView = CollectionViewSource.GetDefaultView(Items);
        ItemsView.Filter = FilterItem;
        var (localInfo, remoteInfo) = GitRunner.MergeInfo(driver?.CurrentPath ?? merged);
        LocalCommit = FormatCommit("本地", localInfo);
        RemoteCommit = FormatCommit("线上", remoteInfo);
        TargetPath = merged;
        BackupRoot = _settings.BackupRootDir;
        SkipNewRows = _settings.SkipNewRows;
        SkipNewColumns = _settings.SkipNewColumns;
        DeleteMissingRows = _settings.DeleteMissingRows;
        DeleteMissingColumns = _settings.DeleteMissingColumns;
        AddNewSheets = _settings.AddNewSheets;
        DeleteMissingSheets = _settings.DeleteMissingSheets;
        ResolveConflicts = _settings.ResolveConflicts;
        UpdateWizardBindings();
    }

    public ObservableCollection<PreviewRow> Items { get; } = [];
    public ICollectionView ItemsView { get; }

    [ObservableProperty] private string _statusText = "正在加载预览...";
    [ObservableProperty] private string _summaryText = "";
    [ObservableProperty] private string _searchText = "";
    [ObservableProperty] private string _currentSheetName = "正在加载...";
    [ObservableProperty] private string _sheetProgressText = "0/0";
    [ObservableProperty] private string _currentSheetStatusText = "";
    [ObservableProperty] private string _localCommit = "";
    [ObservableProperty] private string _remoteCommit = "";
    [ObservableProperty] private string _targetPath = "";
    [ObservableProperty] private string _backupRoot = "";
    [ObservableProperty] private bool _busy;
    [ObservableProperty] private bool _canGenerate;
    [ObservableProperty] private bool _canGoPrevious;
    [ObservableProperty] private bool _canGoNext;
    [ObservableProperty] private bool _canFinishCurrentSheet;
    [ObservableProperty] private bool _canConfirm;
    [ObservableProperty] private bool _skipNewRows;
    [ObservableProperty] private bool _skipNewColumns;
    [ObservableProperty] private bool _deleteMissingRows;
    [ObservableProperty] private bool _deleteMissingColumns;
    [ObservableProperty] private bool _addNewSheets;
    [ObservableProperty] private bool _deleteMissingSheets;
    [ObservableProperty] private bool _resolveConflicts;
    [ObservableProperty] private string _baseSide = "local";
    [ObservableProperty] private PreviewRow? _selectedItem;
    [ObservableProperty] private BackupInfo? _lastBackup;

    partial void OnSearchTextChanged(string value) => ItemsView.Refresh();
    partial void OnBusyChanged(bool value) => UpdateWizardBindings();
    partial void OnSkipNewRowsChanged(bool value) => PersistAndRefresh();
    partial void OnSkipNewColumnsChanged(bool value) => PersistAndRefresh();
    partial void OnDeleteMissingRowsChanged(bool value) => PersistAndRefresh();
    partial void OnDeleteMissingColumnsChanged(bool value) => PersistAndRefresh();
    partial void OnAddNewSheetsChanged(bool value) => PersistAndRefresh();
    partial void OnDeleteMissingSheetsChanged(bool value) => PersistAndRefresh();
    partial void OnResolveConflictsChanged(bool value) => PersistAndRefresh();
    partial void OnBaseSideChanged(string value) => PersistAndRefresh();

    public async Task LoadAsync()
    {
        Busy = true;
        StatusText = "正在读取工作簿...";
        try
        {
            await Task.Run(() =>
            {
                _session?.Dispose();
                _session = new MergeSession(_local, _base, _remote);
            });
            RebuildPreview();
            StatusText = "预览已更新";
        }
        catch (Exception ex)
        {
            StatusText = "加载失败: " + ex.Message;
        }
        finally
        {
            Busy = false;
        }
    }

    [RelayCommand]
    private void ChooseLocal(PreviewRow? row) => SetChoice(row, "local");

    [RelayCommand]
    private void ChooseRemote(PreviewRow? row) => SetChoice(row, "remote");

    [RelayCommand]
    private void PreviousSheet()
    {
        if (_wizard.MovePrevious())
        {
            RefreshCurrentSheet();
        }
    }

    [RelayCommand]
    private void NextSheet()
    {
        if (_wizard.MoveNext())
        {
            RefreshCurrentSheet();
        }
        else if (_wizard.CurrentSheet is { IsCompleted: false })
        {
            StatusText = "请先点击“完成本 Sheet”再继续。";
        }
    }

    [RelayCommand]
    private void FinishCurrentSheet()
    {
        if (_wizard.TryCompleteCurrentSheet())
        {
            StatusText = $"已完成 Sheet：{_wizard.CurrentSheet?.Name}";
            UpdateWizardBindings();
        }
    }

    [RelayCommand]
    private async Task GenerateAsync()
    {
        if (_session is null || !CanGenerate || Busy)
        {
            return;
        }

        Busy = true;
        CanConfirm = false;
        StatusText = "正在生成合并结果...";
        try
        {
            var options = CurrentOptions();
            var choices = CurrentChoices();
            var backupRoot = BackupRoot;
            var result = await Task.Run(() => MergeService.Run(
                _local, _base, _remote, _merged,
                options: options,
                choices: choices,
                backupRoot: backupRoot,
                backupContextPath: _driver?.CurrentPath));
            if (result.ExitCode != 0)
            {
                StatusText = result.Error ?? "合并失败";
                return;
            }

            LastBackup = result.Backup;
            CanConfirm = true;
            StatusText = "合并完成" + (result.Backup is null ? "" : "，备份=" + result.Backup.Dir);
        }
        catch (Exception ex)
        {
            StatusText = "合并失败: " + ex.Message;
        }
        finally
        {
            Busy = false;
        }
    }

    public CompletionResult Confirm()
    {
        if (_driver is not null)
        {
            return GitMergeDriver.WriteBack(_driver);
        }

        return GitCompletion.StageAndCleanup(_merged, _local, _base, _remote);
    }

    public string BuildPathAiPrompt() => MergeAiPromptBuilder.BuildPathPrompt(_local, _base, _remote, _merged);

    public string BuildDetailAiPrompt() => _session is not null
        ? MergeAiPromptBuilder.BuildDetailPrompt(_local, _base, _remote, _merged, _session)
        : MergeAiPromptBuilder.BuildDetailPrompt(
            _local,
            _base,
            _remote,
            _merged,
            _preview ?? new MergePreview
            {
                Items = [],
                ConflictEntries = [],
                Summary = new PreviewSummary(),
                BaseSide = BaseSide,
                Options = [],
            });

    public void DisposeSession() => _session?.Dispose();

    private void PersistAndRefresh()
    {
        if (_session is null)
        {
            return;
        }

        _settings.SkipNewRows = SkipNewRows;
        _settings.SkipNewColumns = SkipNewColumns;
        _settings.DeleteMissingRows = DeleteMissingRows;
        _settings.DeleteMissingColumns = DeleteMissingColumns;
        _settings.AddNewSheets = AddNewSheets;
        _settings.DeleteMissingSheets = DeleteMissingSheets;
        _settings.ResolveConflicts = ResolveConflicts;
        _settings.BackupRootDir = BackupRoot;
        AppSettingsStore.Save(_settings);
        RebuildPreview();
    }

    private void RebuildPreview()
    {
        if (_session is null)
        {
            return;
        }

        _preview = PreviewBuilder.Build(_session, CurrentOptions());
        _wizard.ApplyPreview(_preview, _session.UnionSheetNames());
        _allItems = _preview.Items
            .Select(item => PreviewRow.From(item, _preview.ConflictEntries))
            .ToList();
        RestoreChoices();
        RefreshCurrentSheet();

        var s = _preview.Summary;
        SummaryText = $"基准={(BaseSide == "remote" ? "线上" : "本地")}；新增 {s.New}；删除 {s.Delete}；冲突 {s.Conflict}；信息 {s.Info}；合计 {_allItems.Count}";
        CanConfirm = false;
    }

    private MergeOptions CurrentOptions() => new()
    {
        SkipNewRows = SkipNewRows,
        SkipNewColumns = SkipNewColumns,
        DeleteMissingRows = DeleteMissingRows,
        DeleteMissingColumns = DeleteMissingColumns,
        AddNewSheets = AddNewSheets,
        DeleteMissingSheets = DeleteMissingSheets,
        ResolveConflicts = ResolveConflicts,
        BaseSide = BaseSide,
    };

    private List<MergeChoice> CurrentChoices() => _wizard.Choices.ToList();

    private void SetChoice(PreviewRow? row, string side)
    {
        row ??= SelectedItem;
        if (row is null || !row.CanChoose)
        {
            return;
        }

        var sheet = row.ConflictSheet ?? row.Sheet;
        var key = row.ConflictKey ?? row.Key;
        var kind = row.IsColumn ? ConflictKind.Column : ConflictKind.Row;
        if (_wizard.TrySetChoice(sheet, key, kind, side))
        {
            row.ApplyChoice(side);
            ItemsView.Refresh();
            UpdateWizardBindings();
        }
    }

    private bool FilterItem(object obj)
    {
        if (obj is not PreviewRow row)
        {
            return false;
        }

        if (!string.IsNullOrWhiteSpace(SearchText) &&
            row.Key.IndexOf(SearchText, StringComparison.OrdinalIgnoreCase) < 0 &&
            row.Sheet.IndexOf(SearchText, StringComparison.OrdinalIgnoreCase) < 0)
        {
            return false;
        }

        return true;
    }

    private void RestoreChoices()
    {
        var choices = _wizard.Choices.ToDictionary(choice => choice.ChoiceKey, StringComparer.Ordinal);
        foreach (var row in _allItems.Where(row => row.CanChoose))
        {
            var key = new MergeChoice
            {
                Sheet = row.ConflictSheet ?? row.Sheet,
                Key = row.ConflictKey ?? row.Key,
                Kind = row.IsColumn ? ConflictKind.Column : ConflictKind.Row,
            }.ChoiceKey;
            if (choices.TryGetValue(key, out var choice))
            {
                row.ApplyChoice(choice.Choice);
            }
        }
    }

    private void RefreshCurrentSheet()
    {
        var sheet = _wizard.CurrentSheet?.Name;
        Items.Clear();
        if (sheet is not null)
        {
            foreach (var row in _allItems.Where(row => string.Equals(row.Sheet, sheet, StringComparison.Ordinal)))
            {
                Items.Add(row);
            }
        }

        SelectedItem = null;
        ItemsView.Refresh();
        UpdateWizardBindings();
    }

    private void UpdateWizardBindings()
    {
        var current = _wizard.CurrentSheet;
        CurrentSheetName = current?.Name ?? "无可处理 Sheet";
        SheetProgressText = _wizard.Sheets.Count == 0
            ? "0/0"
            : $"{_wizard.CurrentIndex + 1}/{_wizard.Sheets.Count}";
        CurrentSheetStatusText = current is null
            ? ""
            : current.IsCompleted
                ? "本 Sheet 已完成"
                : current.HasSelectableConflicts ? "本 Sheet 有待处理冲突" : "本 Sheet 无冲突，等待确认";
        CanGoPrevious = _wizard.CurrentIndex > 0;
        CanGoNext = current is not null && current.IsCompleted && _wizard.CurrentIndex < _wizard.Sheets.Count - 1;
        CanFinishCurrentSheet = current is { IsCompleted: false };
        CanGenerate = _wizard.CanGenerate && !Busy;
    }

    private static string FormatCommit(string title, GitCommitInfo? info)
    {
        if (info is null || string.IsNullOrEmpty(info.Hash))
        {
            return title + "：无提交信息";
        }

        return $"{title}：{info.Author}  {info.ShortHash}  {info.Message}";
    }
}

public sealed partial class PreviewRow : ObservableObject
{
    public required string Sheet { get; init; }
    public required string Key { get; init; }
    [ObservableProperty] private string _action = "";
    public required string TypeLabel { get; init; }
    public required string Tag { get; init; }
    public bool CanChoose { get; init; }
    public bool IsColumn { get; init; }
    public string? ConflictSheet { get; init; }
    public string? ConflictKey { get; init; }
    [ObservableProperty] private string? _choice;
    public IReadOnlyList<string> LocalValues { get; init; } = [];
    public IReadOnlyList<string> RemoteValues { get; init; } = [];
    public IReadOnlyList<string> BaseValues { get; init; } = [];
    public IReadOnlyList<string> LocalDisplay => Display(LocalValues);
    public IReadOnlyList<string> RemoteDisplay => Display(RemoteValues);
    public IReadOnlyList<string> BaseDisplay => Display(BaseValues);

    private static IReadOnlyList<string> Display(IReadOnlyList<string> values) =>
        values.Count > 0 ? values : ["（此侧没有该行）"];

    public void ApplyChoice(string side)
    {
        Choice = side == "remote" || side == "线上" ? "线上" : "本地";
        Action = IsColumn
            ? (Choice == "线上" ? "将保留线上列" : "将保留本地列")
            : (Choice == "线上" ? "将保留线上" : "将保留本地");
    }

    public static PreviewRow From(PreviewItem item, IReadOnlyList<MergeChoice> conflicts)
    {
        MergeChoice? choice = item.ConflictIndex is int idx && idx >= 0 && idx < conflicts.Count
            ? conflicts[idx]
            : null;
        return new PreviewRow
        {
            Sheet = item.Sheet,
            Key = item.Key,
            Action = item.Action,
            TypeLabel = item.Tag switch
            {
                PreviewTag.New => "新增",
                PreviewTag.Delete => "删除",
                PreviewTag.Modify => "自动",
                PreviewTag.DeleteConflict => "删除冲突",
                PreviewTag.Conflict => "冲突",
                _ => "信息",
            },
            Tag = item.Tag.ToString(),
            CanChoose = choice is not null,
            IsColumn = choice?.Kind == ConflictKind.Column,
            ConflictSheet = choice?.Sheet,
            ConflictKey = choice?.Key,
            Choice = choice is null ? null : choice.Choice == "remote" ? "线上" : "本地",
            LocalValues = item.LocalValues,
            RemoteValues = item.RemoteValues,
            BaseValues = item.BaseValues,
        };
    }
}
