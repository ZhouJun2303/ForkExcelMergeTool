using ExcelMergeFork.Core.Git;
using ExcelMergeFork.Core.Merge;

namespace ExcelMergeFork.Api;

public sealed class MergeAiPanel
{
    private readonly MergeWorkSession _work;
    private readonly List<AiFileRow> _files = [];

    public MergeAiPanel(MergeWorkSession work) => _work = work;

    public string StatusText { get; set; } = "";
    public string RepoText { get; private set; } = "";
    public string SummaryText { get; private set; } = "没有可交给 AI 的 Excel";

    public void Refresh()
    {
        StatusText = "正在查找冲突 Excel...";
        var repo = ResolveRepoRoot();
        var scan = repo is null ? new RepoConflictScan() : RepoConflictExcelFinder.Scan(repo);
        Apply(repo, scan);
    }

    public void SetAll(bool value)
    {
        foreach (var file in _files)
        {
            file.IsChecked = value;
        }

        UpdateSummary();
    }

    public void SetChecked(string? path, bool value)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new InvalidOperationException("缺少文件路径");
        }

        var file = _files.FirstOrDefault(item => string.Equals(item.FullPath, path, StringComparison.OrdinalIgnoreCase));
        if (file is null)
        {
            throw new InvalidOperationException("列表里没有这个文件");
        }

        file.IsChecked = value;
        UpdateSummary();
    }

    public string BuildPrompt(bool detail)
    {
        var selected = _files.Where(file => file.IsChecked).ToList();
        if (selected.Count == 0)
        {
            StatusText = "请先勾选要交给 AI 的 Excel";
            throw new InvalidOperationException(StatusText);
        }

        StatusText = detail ? "正在读取勾选的 Excel..." : "正在整理勾选的路径...";
        var currentPreview = detail && _work.Preview is not null ? CurrentPreviewForAi() : null;
        var missingPreview = detail && _work.Preview is null;
        var books = selected.Select(file => ToPromptBook(file, detail, currentPreview, missingPreview)).ToList();
        StatusText = "已整理 " + books.Count + " 个 Excel";
        return detail
            ? MergeAiPromptBuilder.BuildBatchDetailPrompt(books)
            : MergeAiPromptBuilder.BuildBatchPathPrompt(books);
    }

    public object Snapshot() => new
    {
        RepoText,
        SummaryText,
        StatusText,
        Files = _files.Select(file => new
        {
            file.RelativePath,
            file.FullPath,
            file.StatusLabel,
            file.SidesText,
            file.ExtraPathText,
            file.IsCurrent,
            file.IsChecked,
        }).ToList(),
    };

    private void Apply(string? repo, RepoConflictScan scan)
    {
        var previous = _files.ToDictionary(file => file.FullPath, file => file.IsChecked, StringComparer.OrdinalIgnoreCase);
        var currentFull = repo is null ? null : CurrentWorktreePath(repo);
        var seenCurrent = false;
        var picks = new List<AiFileRow>();
        foreach (var file in scan.Files)
        {
            var isCurrent = currentFull is not null && SamePath(file.WorktreePath, currentFull);
            if (isCurrent)
            {
                seenCurrent = true;
            }

            picks.Add(Create(file.RelativePath, file.WorktreePath, isCurrent, file.HasBothSides, isCurrent, repo, file, previous));
        }

        if (!seenCurrent)
        {
            var full = currentFull ?? SafeFull(_work.RequestedMerged ?? _work.LocalSource);
            var relative = repo is not null && currentFull is not null
                ? Path.GetRelativePath(repo, currentFull).Replace('\\', '/')
                : Path.GetFileName(full);
            picks.Insert(0, Create(relative, full, true, true, true, repo, null, previous));
        }

        picks = picks
            .OrderBy(file => file.IsCurrent ? 0 : 1)
            .ThenBy(file => file.RelativePath, StringComparer.OrdinalIgnoreCase)
            .ToList();
        _files.Clear();
        _files.AddRange(picks);
        RepoText = repo is null ? "未找到 Git 仓库，下面只列出当前正在合并的文件。" : "仓库：" + repo;
        StatusText = scan.Error
            ?? (repo is null
                ? "当前文件不在 Git 仓库中"
                : scan.Files.Count == 0
                    ? "当前仓库没有其他未合并的 Excel"
                    : "已找到 " + scan.Files.Count + " 个未合并的 Excel");
        UpdateSummary();
    }

    private AiFileRow Create(
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
        return new AiFileRow
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

    private AiPromptBook ToPromptBook(AiFileRow file, bool detail, MergePreview? currentPreview, bool missingPreview)
    {
        if (file.UseLaunchPaths)
        {
            return new AiPromptBook
            {
                Label = file.RelativePath,
                WorktreePath = file.FullPath,
                Local = _work.LocalSource,
                Base = _work.BaseSource,
                Remote = _work.RemoteSource,
                Merged = _work.RequestedMerged ?? "",
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
        var choices = _work.Wizard.Choices.ToDictionary(choice => choice.ChoiceKey, StringComparer.Ordinal);
        if (_work.Preview is null)
        {
            return new MergePreview
            {
                Items = [],
                ConflictEntries = _work.Wizard.Choices.ToList(),
                Summary = new PreviewSummary(),
                BaseSide = _work.Options.BaseSide,
                Options = [],
            };
        }

        var preview = _work.Preview;
        var entries = preview.ConflictEntries.Select(entry => new MergeChoice
        {
            Sheet = entry.Sheet,
            Key = entry.Key,
            Kind = entry.Kind,
            Choice = choices.TryGetValue(entry.ChoiceKey, out var chosen) ? chosen.Choice : entry.Choice,
            AutoType = entry.AutoType,
        }).ToList();
        var items = preview.Items.Select(item =>
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
            Summary = preview.Summary,
            BaseSide = preview.BaseSide,
            Options = preview.Options,
            ElapsedMs = preview.ElapsedMs,
            SheetCount = preview.SheetCount,
        };
    }

    private void UpdateSummary()
    {
        var checkedCount = _files.Count(file => file.IsChecked);
        SummaryText = _files.Count == 0
            ? "没有可交给 AI 的 Excel"
            : "共 " + _files.Count + " 个，已勾选 " + checkedCount + " 个";
    }

    private string? ResolveRepoRoot()
    {
        foreach (var candidate in new[]
        {
            _work.Driver?.RepoPath,
            _work.Driver?.CurrentPath,
            _work.RequestedMerged,
            _work.LocalSource,
            _work.RemoteSource,
            _work.BaseSource,
        })
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
        foreach (var candidate in new[]
        {
            _work.Driver?.CurrentPath,
            _work.RequestedMerged,
            _work.LocalSource,
            _work.RemoteSource,
            _work.BaseSource,
        })
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
        "本地：" + SafeFull(_work.LocalSource) + Environment.NewLine +
        "基准：" + SafeFull(_work.BaseSource) + Environment.NewLine +
        "线上：" + SafeFull(_work.RemoteSource) + Environment.NewLine +
        "输出：" + SafeFull(_work.RequestedMerged ?? "");

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
            return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
        }
    }

    private static string SafeFull(string path)
    {
        try
        {
            return string.IsNullOrWhiteSpace(path) ? "" : Path.GetFullPath(path);
        }
        catch
        {
            return path;
        }
    }

    private sealed class AiFileRow
    {
        public required string RelativePath { get; init; }
        public required string FullPath { get; init; }
        public required string StatusLabel { get; init; }
        public string SidesText { get; init; } = "";
        public string ExtraPathText { get; init; } = "";
        public bool IsCurrent { get; init; }
        public bool HasBothSides { get; init; }
        public bool UseLaunchPaths { get; init; }
        public string? RepoRoot { get; init; }
        public RepoConflictExcel? Source { get; init; }
        public bool IsChecked { get; set; }
    }
}
