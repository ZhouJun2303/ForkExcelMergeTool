using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
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
    public ObservableCollection<AiConflictPick> AiFiles { get; } = [];
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
    [ObservableProperty] private bool _canAttemptGenerate;
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
    [ObservableProperty] private string _aiRepoText = "";
    [ObservableProperty] private string _aiSummaryText = "正在查找冲突 Excel...";
    [ObservableProperty] private string _aiStatusText = "正在查找冲突 Excel...";
    [ObservableProperty] private bool _canUseAi = true;

    partial void OnSearchTextChanged(string value) => ItemsView.Refresh();
    partial void OnBusyChanged(bool value)
    {
        CanUseAi = !value;
        UpdateWizardBindings();
    }
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
            try
            {
                await LoadAiFilesAsync();
            }
            catch (Exception ex)
            {
                AiStatusText = "查找冲突 Excel 失败：" + ex.Message;
            }

            Busy = false;
        }
    }

    [RelayCommand]
    private async Task RefreshAiFilesAsync()
    {
        if (Busy)
        {
            return;
        }

        Busy = true;
        try
        {
            await LoadAiFilesAsync();
        }
        catch (Exception ex)
        {
            AiStatusText = "查找冲突 Excel 失败：" + ex.Message;
        }
        finally
        {
            Busy = false;
        }
    }

    [RelayCommand]
    private void CheckAllAiFiles() => SetAllAiChecks(true);

    [RelayCommand]
    private void UncheckAllAiFiles() => SetAllAiChecks(false);

    public async Task<string?> BuildCheckedPathPromptAsync() => await BuildCheckedPromptAsync(detail: false);

    public async Task<string?> BuildCheckedDetailPromptAsync() => await BuildCheckedPromptAsync(detail: true);

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
            StatusText = "请先点击“确认本 Sheet”再继续。";
        }
    }

    [RelayCommand]
    private void FinishCurrentSheet()
    {
        var name = _wizard.CurrentSheet?.Name;
        if (!_wizard.TryCompleteCurrentSheet())
        {
            return;
        }

        StatusText = $"已确认 Sheet：{name}";
        if (_wizard.MoveNext())
        {
            RefreshCurrentSheet();
            return;
        }

        UpdateWizardBindings();
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
                ? "本 Sheet 已确认"
                : current.HasSelectableConflicts ? "本 Sheet 有待处理冲突" : "本 Sheet 无冲突，等待确认";
        CanGoPrevious = _wizard.CurrentIndex > 0;
        CanGoNext = current is not null && current.IsCompleted && _wizard.CurrentIndex < _wizard.Sheets.Count - 1;
        CanFinishCurrentSheet = current is { IsCompleted: false };
        CanGenerate = _wizard.CanGenerate && !Busy;
        CanAttemptGenerate = _session is not null && !Busy;
    }

    public bool AllSheetsConfirmed => _wizard.CanGenerate;

    public IReadOnlyList<string> UnconfirmedSheetNames => _wizard.UnconfirmedSheetNames;

    private async Task LoadAiFilesAsync()
    {
        AiStatusText = "正在查找冲突 Excel...";
        var repo = ResolveRepoRoot();
        var scan = repo is null
            ? new RepoConflictScan()
            : await Task.Run(() => RepoConflictExcelFinder.Scan(repo));
        ApplyAiFiles(repo, scan);
    }

    private async Task<string?> BuildCheckedPromptAsync(bool detail)
    {
        var selected = AiFiles.Where(file => file.IsChecked).ToList();
        if (selected.Count == 0)
        {
            AiStatusText = "请先勾选要交给 AI 的 Excel";
            return null;
        }

        Busy = true;
        AiStatusText = detail ? "正在读取勾选的 Excel..." : "正在整理勾选的路径...";
        try
        {
            var currentPreview = detail && _preview is not null ? CurrentPreviewForAi() : null;
            var missingPreview = detail && _preview is null;
            var books = await Task.Run(() => selected.Select(file => ToPromptBook(file, detail, currentPreview, missingPreview)).ToList());
            AiStatusText = $"已整理 {books.Count} 个 Excel";
            return detail
                ? MergeAiPromptBuilder.BuildBatchDetailPrompt(books)
                : MergeAiPromptBuilder.BuildBatchPathPrompt(books);
        }
        catch (Exception ex)
        {
            AiStatusText = "整理失败：" + ex.Message;
            return null;
        }
        finally
        {
            Busy = false;
        }
    }

    private void ApplyAiFiles(string? repo, RepoConflictScan scan)
    {
        var previous = AiFiles.ToDictionary(file => file.FullPath, file => file.IsChecked, StringComparer.OrdinalIgnoreCase);
        var currentFull = repo is null ? null : CurrentWorktreePath(repo);
        var seenCurrent = false;
        var picks = new List<AiConflictPick>();
        foreach (var file in scan.Files)
        {
            var isCurrent = currentFull is not null && SamePath(file.WorktreePath, currentFull);
            if (isCurrent)
            {
                seenCurrent = true;
            }

            picks.Add(CreatePick(file.RelativePath, file.WorktreePath, isCurrent, file.HasBothSides, isCurrent, repo, file, previous));
        }

        if (!seenCurrent)
        {
            var full = currentFull ?? SafeFullPath(_merged);
            var relative = repo is not null && currentFull is not null
                ? Path.GetRelativePath(repo, currentFull).Replace('\\', '/')
                : Path.GetFileName(full);
            picks.Insert(0, CreatePick(relative, full, true, true, true, repo, null, previous));
        }

        picks = picks
            .OrderBy(file => file.IsCurrent ? 0 : 1)
            .ThenBy(file => file.RelativePath, StringComparer.OrdinalIgnoreCase)
            .ToList();
        AiFiles.Clear();
        foreach (var pick in picks)
        {
            pick.CheckedChanged += (_, _) => UpdateAiSummary();
            AiFiles.Add(pick);
        }

        AiRepoText = repo is null ? "未找到 Git 仓库，下面只列出当前正在合并的文件。" : "仓库：" + repo;
        AiStatusText = scan.Error
            ?? (repo is null
                ? "当前文件不在 Git 仓库中"
                : scan.Files.Count == 0
                    ? "当前仓库没有其他未合并的 Excel"
                    : $"已找到 {scan.Files.Count} 个未合并的 Excel");
        UpdateAiSummary();
    }

    private AiConflictPick CreatePick(
        string relative,
        string full,
        bool isCurrent,
        bool hasBothSides,
        bool useLaunchPaths,
        string? repo,
        RepoConflictExcel? source,
        Dictionary<string, bool> previous)
    {
        var status = isCurrent && hasBothSides && source is not null
            ? "当前正在合并 · 双方都有修改"
            : isCurrent
                ? "当前正在合并"
                : hasBothSides
                    ? "双方都有修改"
                    : "冲突文件不完整";
        var checkedByDefault = isCurrent || hasBothSides;
        return new AiConflictPick
        {
            RelativePath = relative,
            FullPath = full,
            StatusLabel = status,
            SidesText = SidesText(source, useLaunchPaths),
            ExtraPathText = useLaunchPaths ? LaunchPathText() : "",
            IsCurrent = isCurrent,
            HasBothSides = hasBothSides,
            UseLaunchPaths = useLaunchPaths,
            RepoRoot = repo,
            Source = source,
            IsChecked = previous.TryGetValue(full, out var wasChecked) ? wasChecked : checkedByDefault,
        };
    }

    private AiPromptBook ToPromptBook(AiConflictPick file, bool detail, MergePreview? currentPreview, bool missingPreview)
    {
        if (file.UseLaunchPaths)
        {
            return new AiPromptBook
            {
                Label = file.RelativePath,
                WorktreePath = file.FullPath,
                Local = _local,
                Base = _base,
                Remote = _remote,
                Merged = _merged,
                Preview = detail ? currentPreview : null,
                ReadError = missingPreview ? "当前预览尚未加载" : null,
            };
        }

        if (file.Source is null || string.IsNullOrWhiteSpace(file.RepoRoot))
        {
            return new AiPromptBook
            {
                Label = file.RelativePath,
                WorktreePath = file.FullPath,
                ReadError = "没有仓库阶段信息",
            };
        }

        var stageRoot = AiMergeBatch.StageRootFor(file.RepoRoot);
        return detail
            ? AiMergeBatch.MaterializeWithPreview(file.RepoRoot, file.Source, stageRoot)
            : AiMergeBatch.MaterializePaths(file.RepoRoot, file.Source, stageRoot);
    }

    private MergePreview CurrentPreviewForAi()
    {
        if (_preview is null)
        {
            return new MergePreview
            {
                Items = [],
                ConflictEntries = CurrentChoices(),
                Summary = new PreviewSummary(),
                BaseSide = BaseSide,
                Options = [],
            };
        }

        var choices = CurrentChoices().ToDictionary(choice => choice.ChoiceKey, StringComparer.Ordinal);
        var entries = _preview.ConflictEntries.Select(entry => new MergeChoice
        {
            Sheet = entry.Sheet,
            Key = entry.Key,
            Kind = entry.Kind,
            Choice = choices.TryGetValue(entry.ChoiceKey, out var chosen) ? chosen.Choice : entry.Choice,
            AutoType = entry.AutoType,
        }).ToList();
        var items = _preview.Items.Select(item =>
        {
            var action = item.Action;
            if (item.ConflictIndex is int index && index >= 0 && index < entries.Count)
            {
                var entry = entries[index];
                var remote = entry.Choice == "remote";
                action = entry.Kind == ConflictKind.Column
                    ? remote ? "将保留线上列" : "将保留本地列"
                    : remote ? "将保留线上" : "将保留本地";
            }

            return new PreviewItem
            {
                Sheet = item.Sheet,
                Key = item.Key,
                Action = action,
                Tag = item.Tag,
                ConflictIndex = item.ConflictIndex,
                LocalValues = item.LocalValues,
                RemoteValues = item.RemoteValues,
                BaseValues = item.BaseValues,
            };
        }).ToList();
        return new MergePreview
        {
            Items = items,
            ConflictEntries = entries,
            Summary = _preview.Summary,
            BaseSide = _preview.BaseSide,
            Options = _preview.Options,
            ElapsedMs = _preview.ElapsedMs,
            SheetCount = _preview.SheetCount,
        };
    }

    private void SetAllAiChecks(bool value)
    {
        foreach (var file in AiFiles)
        {
            file.IsChecked = value;
        }

        UpdateAiSummary();
    }

    private void UpdateAiSummary()
    {
        var checkedCount = AiFiles.Count(file => file.IsChecked);
        AiSummaryText = AiFiles.Count == 0
            ? "没有可交给 AI 的 Excel"
            : $"共 {AiFiles.Count} 个，已勾选 {checkedCount} 个";
    }

    private string? ResolveRepoRoot()
    {
        foreach (var candidate in new[] { _driver?.RepoPath, _driver?.CurrentPath, _merged, _local, _remote, _base })
        {
            var repo = GitRunner.FindWorktreeRoot(candidate);
            if (!string.IsNullOrWhiteSpace(repo))
            {
                return repo;
            }
        }

        return null;
    }

    private string? CurrentWorktreePath(string repo)
    {
        foreach (var candidate in new[] { _driver?.CurrentPath, _merged, _local, _remote, _base })
        {
            var inside = PathInsideRepo(candidate, repo);
            if (inside is not null)
            {
                return inside;
            }
        }

        return null;
    }

    private string LaunchPathText() =>
        "本地：" + SafeFullPath(_local) + Environment.NewLine +
        "基准：" + SafeFullPath(_base) + Environment.NewLine +
        "线上：" + SafeFullPath(_remote) + Environment.NewLine +
        "输出：" + SafeFullPath(_merged);

    private static string SidesText(RepoConflictExcel? source, bool launch)
    {
        if (launch || source is null)
        {
            return "本地 · 基准 · 线上";
        }

        return string.Join(" · ",
            source.LocalBlob is null ? "无本地" : "本地",
            source.BaseBlob is null ? "无基准" : "基准",
            source.RemoteBlob is null ? "无线上" : "线上");
    }

    private static string? PathInsideRepo(string? path, string repo)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            return null;
        }

        try
        {
            var full = Path.GetFullPath(path);
            var root = Path.GetFullPath(repo).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            var prefix = root + Path.DirectorySeparatorChar;
            return full.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) ? full : null;
        }
        catch
        {
            return null;
        }
    }

    private static bool SamePath(string left, string right)
    {
        try
        {
            return string.Equals(Path.GetFullPath(left), Path.GetFullPath(right), StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    private static string SafeFullPath(string path)
    {
        try
        {
            return Path.GetFullPath(path);
        }
        catch
        {
            return path;
        }
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
