using ExcelMergeFork.Core.Excel;
using ExcelMergeFork.Core.Git;
using ExcelMergeFork.Core.Merge;
using ExcelMergeFork.Core.Settings;

namespace ExcelMergeFork.Api;

public sealed class MergeWorkSession : IDisposable
{
    private bool _disposed;

    private MergeWorkSession(
        string id,
        string ownedDir,
        string localSource,
        string baseSource,
        string remoteSource,
        string local,
        string basePath,
        string remote,
        string? requestedMerged,
        MergeOptions options,
        MergeSession session)
    {
        Id = id;
        OwnedDir = ownedDir;
        LocalSource = localSource;
        BaseSource = baseSource;
        RemoteSource = remoteSource;
        Local = local;
        Base = basePath;
        Remote = remote;
        RequestedMerged = requestedMerged;
        Options = options;
        Session = session;
        Ai = new MergeAiPanel(this);
        TouchedUtc = DateTime.UtcNow;
    }

    public object Gate { get; } = new();
    public string Id { get; }
    public string OwnedDir { get; }
    public string LocalSource { get; }
    public string BaseSource { get; }
    public string RemoteSource { get; }
    public string Local { get; }
    public string Base { get; }
    public string Remote { get; }
    public string? RequestedMerged { get; private set; }
    public MergeOptions Options { get; private set; }
    public MergeSession Session { get; }
    public MergeWizardState Wizard { get; } = new();
    public MergePreview? Preview { get; private set; }
    public string? GeneratedPath { get; private set; }
    public string? BackupDir { get; private set; }
    public bool CanConfirm { get; private set; }
    public GitDriverRequest? Driver { get; private set; }
    public string LocalCommit { get; private set; } = "本地：无提交信息";
    public string RemoteCommit { get; private set; } = "线上：无提交信息";
    public MergeAiPanel Ai { get; }
    public string Status { get; private set; } = "正在读取工作簿...";
    public DateTime TouchedUtc { get; private set; }

    public static MergeWorkSession Open(string local, string basePath, string remote, string? requestedMerged, MergeOptions options)
    {
        var id = Guid.NewGuid().ToString("N");
        var dir = Path.Combine(RepoLayout.WorkRoot, "merge-" + id);
        Directory.CreateDirectory(dir);
        MergeSession? session = null;
        try
        {
            var localCopy = Path.Combine(dir, "local" + Path.GetExtension(local));
            var baseCopy = Path.Combine(dir, "base" + Path.GetExtension(basePath));
            var remoteCopy = Path.Combine(dir, "remote" + Path.GetExtension(remote));
            File.Copy(local, localCopy, true);
            File.Copy(basePath, baseCopy, true);
            File.Copy(remote, remoteCopy, true);
            session = new MergeSession(localCopy, baseCopy, remoteCopy);
            var work = new MergeWorkSession(
                id,
                dir,
                local,
                basePath,
                remote,
                localCopy,
                baseCopy,
                remoteCopy,
                requestedMerged,
                options,
                session);
            work.Rebuild();
            work.LoadCommits();
            try
            {
                work.Ai.Refresh();
            }
            catch (Exception ex)
            {
                work.Ai.StatusText = "查找冲突 Excel 失败：" + ex.Message;
            }

            return work;
        }
        catch
        {
            session?.Dispose();
            TryDelete(dir);
            throw;
        }
    }

    public static MergeOptions FromSettings()
    {
        var settings = AppSettingsStore.Load();
        return new MergeOptions
        {
            SkipNewRows = settings.SkipNewRows,
            SkipNewColumns = settings.SkipNewColumns,
            DeleteMissingRows = settings.DeleteMissingRows,
            DeleteMissingColumns = settings.DeleteMissingColumns,
            AddNewSheets = settings.AddNewSheets,
            DeleteMissingSheets = settings.DeleteMissingSheets,
            ResolveConflicts = settings.ResolveConflicts,
            BaseSide = "local",
        };
    }

    public static MergeOptions ResolveOptions(MergeOptions current, MergeOptionsBody? body, string? mode)
    {
        var letter = (mode ?? body?.Mode)?.Trim();
        var side = NormalizeSide(body?.BaseSide ?? current.BaseSide);
        if (!string.IsNullOrEmpty(letter))
        {
            var upper = letter.ToUpperInvariant();
            if (upper is not ("A" or "B" or "C" or "D" or "E"))
            {
                throw new InvalidOperationException("模式只支持 A–E");
            }

            return MergeOptions.ForMode(upper, side);
        }

        if (body is null)
        {
            return current;
        }

        return new MergeOptions
        {
            SkipNewRows = body.SkipNewRows,
            SkipNewColumns = body.SkipNewColumns,
            DeleteMissingRows = body.DeleteMissingRows,
            DeleteMissingColumns = body.DeleteMissingColumns,
            AddNewSheets = body.AddNewSheets,
            DeleteMissingSheets = body.DeleteMissingSheets,
            ResolveConflicts = body.ResolveConflicts,
            BaseSide = side,
        };
    }

    public void AttachDriver(GitDriverRequest driver)
    {
        Driver = driver;
        LoadCommits();
    }

    public void ApplyOptions(MergeOptions options)
    {
        Options = options;
        CanConfirm = false;
        Rebuild();
        Status = "预览已更新";
    }

    public void Choose(string? sheet, string? key, string? kind, string? choice)
    {
        if (string.IsNullOrWhiteSpace(sheet) || string.IsNullOrWhiteSpace(key))
        {
            throw new InvalidOperationException("缺少工作表或主键");
        }

        if (!TryParseKind(kind, out var parsed))
        {
            throw new InvalidOperationException("冲突类型无效");
        }

        var side = NormalizeChoice(choice);
        if (!Wizard.TrySetChoice(sheet, key, parsed, side))
        {
            throw new InvalidOperationException("没有这条冲突");
        }

        Status = side == "remote" ? "已选择线上" : "已选择本地";
    }

    public void SelectSheet(int index)
    {
        if (!Wizard.TrySelect(index))
        {
            throw new InvalidOperationException("工作表序号无效");
        }

        Status = "当前表：" + (Wizard.CurrentSheet?.Name ?? "");
    }

    public void CompleteCurrent()
    {
        var name = Wizard.CurrentSheet?.Name ?? throw new InvalidOperationException("没有需要确认的工作表");
        if (!Wizard.TryCompleteCurrentSheet())
        {
            throw new InvalidOperationException("无法确认当前工作表");
        }

        Status = "已确认 " + name;
        Wizard.MoveNext();
    }

    public void PreviousSheet()
    {
        if (!Wizard.MovePrevious())
        {
            throw new InvalidOperationException("已经是第一个工作表");
        }
    }

    public void NextSheet()
    {
        if (Wizard.CurrentSheet is { IsCompleted: false })
        {
            throw new InvalidOperationException("请先确认本工作表");
        }

        if (!Wizard.MoveNext())
        {
            throw new InvalidOperationException("已经是最后一个工作表");
        }
    }

    public void Generate(string? merged, bool createBackup)
    {
        if (!Wizard.CanGenerate)
        {
            var pending = string.Join("、", Wizard.UnconfirmedSheetNames);
            throw new InvalidOperationException(pending.Length == 0 ? "还有工作表未确认" : "还有工作表未确认：" + pending);
        }

        var target = RepoLayout.ResolveOutput(
            string.IsNullOrWhiteSpace(merged) ? RequestedMerged : merged,
            Path.Combine(OwnedDir, "merged.xlsx"),
            Local,
            Base,
            Remote,
            LocalSource,
            BaseSource,
            RemoteSource);
        RequestedMerged = target;
        var result = MergeService.Run(
            Local,
            Base,
            Remote,
            target,
            options: Options,
            choices: Wizard.Choices.ToList(),
            backupRoot: createBackup ? EmptyToNull(AppSettingsStore.Load().BackupRootDir) : null,
            backupContextPath: Driver?.CurrentPath ?? target,
            createBackup: createBackup);
        if (result.ExitCode != 0)
        {
            Status = result.Error ?? "合并失败";
            throw new InvalidOperationException(Status);
        }

        GeneratedPath = target;
        BackupDir = result.Backup?.Dir;
        CanConfirm = true;
        Status = "合并完成" + (result.Backup is null ? "" : "，备份=" + result.Backup.Dir);
    }

    public CompletionResult Confirm()
    {
        if (!CanConfirm || string.IsNullOrWhiteSpace(GeneratedPath))
        {
            throw new InvalidOperationException("请先生成合并结果");
        }

        if (Driver is not null)
        {
            var written = GitMergeDriver.WriteBack(Driver);
            if (written.Success && ExcelMergeHost.Desktop is not null)
            {
                ExcelMergeHost.Desktop.WriteBackSucceeded = true;
            }

            return written;
        }

        return GitCompletion.StageAndCleanup(GeneratedPath, LocalSource, BaseSource, RemoteSource);
    }

    public object Snapshot(PageQuery page)
    {
        Touch();
        var preview = Preview ?? throw new InvalidOperationException("预览还没有生成");
        var choices = Wizard.Choices.ToDictionary(choice => choice.ChoiceKey, StringComparer.Ordinal);
        var current = Wizard.CurrentSheet?.Name;
        IEnumerable<PreviewItem> items = preview.Items;
        if (current is not null)
        {
            items = items.Where(item => string.Equals(item.Sheet, current, StringComparison.Ordinal));
        }

        if (!string.IsNullOrWhiteSpace(page.Q))
        {
            items = items.Where(item =>
                item.Key.Contains(page.Q, StringComparison.OrdinalIgnoreCase) ||
                item.Sheet.Contains(page.Q, StringComparison.OrdinalIgnoreCase) ||
                item.Action.Contains(page.Q, StringComparison.OrdinalIgnoreCase));
        }

        var filtered = items.ToList();
        var rows = filtered.Skip(page.Offset).Take(page.Limit).Select(item => ToItem(item, preview, choices)).ToList();
        var summary = preview.Summary;
        var sideLabel = Options.BaseSide == "remote" ? "线上" : "本地";
        return new
        {
            Id,
            Status,
            LocalCommit,
            RemoteCommit,
            TargetPath = RequestedMerged ?? "",
            CanConfirm,
            SummaryText = $"基准={sideLabel}；新增 {summary.New}；删除 {summary.Delete}；冲突 {summary.Conflict}；信息 {summary.Info}；合计 {preview.Items.Count}",
            Paths = new { Local = LocalSource, Base = BaseSource, Remote = RemoteSource, Merged = RequestedMerged },
            Names = new
            {
                Local = Path.GetFileName(LocalSource),
                Base = Path.GetFileName(BaseSource),
                Remote = Path.GetFileName(RemoteSource),
            },
            Options = new
            {
                Options.SkipNewRows,
                Options.SkipNewColumns,
                Options.DeleteMissingRows,
                Options.DeleteMissingColumns,
                Options.AddNewSheets,
                Options.DeleteMissingSheets,
                Options.ResolveConflicts,
                Options.BaseSide,
                Letters = Options.ToLetterSet().OrderBy(letter => letter, StringComparer.Ordinal).ToArray(),
            },
            Summary = new
            {
                summary.New,
                summary.Delete,
                summary.Conflict,
                summary.Info,
                preview.ElapsedMs,
                preview.SheetCount,
            },
            Wizard = new
            {
                Index = Wizard.CurrentIndex,
                Current = current,
                Count = Wizard.Sheets.Count,
                CanPrevious = Wizard.CurrentIndex > 0,
                CanNext = Wizard.CurrentSheet is { IsCompleted: true } && Wizard.CurrentIndex < Wizard.Sheets.Count - 1,
                CanFinish = Wizard.CurrentSheet is { IsCompleted: false },
                CanGenerate = Wizard.CanGenerate,
                SheetStatus = Wizard.CurrentSheet is null
                    ? ""
                    : Wizard.CurrentSheet.IsCompleted
                        ? "本 Sheet 已确认"
                        : Wizard.CurrentSheet.HasSelectableConflicts ? "本 Sheet 有待处理冲突" : "本 Sheet 无冲突，等待确认",
                Progress = Wizard.Sheets.Count == 0 ? "0/0" : (Wizard.CurrentIndex + 1) + "/" + Wizard.Sheets.Count,
                Pending = Wizard.UnconfirmedSheetNames,
                Sheets = Wizard.Sheets.Select((sheet, index) => new
                {
                    Index = index,
                    sheet.Name,
                    Completed = sheet.IsCompleted,
                    HasConflicts = sheet.HasSelectableConflicts,
                    Current = index == Wizard.CurrentIndex,
                }).ToList(),
            },
            ItemTotal = filtered.Count,
            page.Offset,
            page.Limit,
            Items = rows,
            GeneratedPath,
            BackupDir,
            DownloadUrl = GeneratedPath is null ? null : "/api/merge/sessions/" + Id + "/download",
            Ai = Ai.Snapshot(),
        };
    }

    private void LoadCommits()
    {
        try
        {
            var context = Driver?.CurrentPath ?? RequestedMerged ?? LocalSource;
            var (localInfo, remoteInfo) = GitRunner.MergeInfo(context);
            LocalCommit = FormatCommit("本地", localInfo);
            RemoteCommit = FormatCommit("线上", remoteInfo);
        }
        catch (Exception ex)
        {
            LocalCommit = "本地：无提交信息";
            RemoteCommit = "线上：" + ex.Message;
        }
    }

    private static string FormatCommit(string title, GitCommitInfo? info)
    {
        if (info is null || string.IsNullOrEmpty(info.Hash))
        {
            return title + "：无提交信息";
        }

        return title + "：" + info.Author + "  " + info.ShortHash + "  " + info.Message;
    }

    public void Touch() => TouchedUtc = DateTime.UtcNow;

    public void Dispose()
    {
        lock (Gate)
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            Session.Dispose();
        }

        TryDelete(OwnedDir);
    }

    private void Rebuild()
    {
        Preview = PreviewBuilder.Build(Session, Options);
        Wizard.ApplyPreview(Preview, Session.UnionSheetNames());
        if (string.IsNullOrEmpty(Status) || Status.StartsWith("正在", StringComparison.Ordinal))
        {
            Status = "预览已更新";
        }
    }

    private static object ToItem(PreviewItem item, MergePreview preview, Dictionary<string, MergeChoice> choices)
    {
        MergeChoice? entry = item.ConflictIndex is int index && index >= 0 && index < preview.ConflictEntries.Count
            ? preview.ConflictEntries[index]
            : null;
        string? choice = null;
        if (entry is not null && choices.TryGetValue(entry.ChoiceKey, out var saved))
        {
            choice = saved.Choice;
        }

        var action = item.Action;
        if (entry is not null && choice is not null)
        {
            var remote = choice == "remote";
            action = entry.Kind == ConflictKind.Column
                ? remote ? "将保留线上列" : "将保留本地列"
                : remote ? "将保留线上" : "将保留本地";
        }

        return new
        {
            item.Sheet,
            item.Key,
            Action = action,
            Tag = item.Tag.ToString(),
            TypeLabel = item.Tag switch
            {
                PreviewTag.New => "新增",
                PreviewTag.Delete => "删除",
                PreviewTag.Modify => "自动",
                PreviewTag.DeleteConflict => "删除冲突",
                PreviewTag.Conflict => "冲突",
                _ => "信息",
            },
            CanChoose = entry is not null,
            IsColumn = entry?.Kind == ConflictKind.Column,
            ConflictSheet = entry?.Sheet,
            ConflictKey = entry?.Key,
            Kind = entry?.Kind.ToString(),
            Choice = choice,
            BaseValues = CapLines(item.BaseValues),
            LocalValues = CapLines(item.LocalValues),
            RemoteValues = CapLines(item.RemoteValues),
        };
    }

    private static IReadOnlyList<string> CapLines(IReadOnlyList<string> values)
    {
        const int maxLines = 80;
        const int maxChars = 2000;
        var taken = values.Take(maxLines).Select(value => value.Length <= maxChars ? value : value[..maxChars] + "…").ToList();
        if (values.Count > maxLines)
        {
            taken.Add("… 还有 " + (values.Count - maxLines) + " 项");
        }

        return taken;
    }

    private static string NormalizeSide(string? side)
    {
        if (string.Equals(side, "remote", StringComparison.OrdinalIgnoreCase) || side == "线上")
        {
            return "remote";
        }

        if (string.IsNullOrWhiteSpace(side) ||
            string.Equals(side, "local", StringComparison.OrdinalIgnoreCase) ||
            side == "本地")
        {
            return "local";
        }

        throw new InvalidOperationException("基准侧只支持本地或线上");
    }

    private static string NormalizeChoice(string? choice)
    {
        if (string.Equals(choice, "remote", StringComparison.OrdinalIgnoreCase) || choice == "线上")
        {
            return "remote";
        }

        if (string.Equals(choice, "local", StringComparison.OrdinalIgnoreCase) || choice == "本地")
        {
            return "local";
        }

        throw new InvalidOperationException("选择只支持本地或线上");
    }

    private static bool TryParseKind(string? kind, out ConflictKind value)
    {
        if (string.IsNullOrWhiteSpace(kind) || string.Equals(kind, "row", StringComparison.OrdinalIgnoreCase))
        {
            value = ConflictKind.Row;
            return true;
        }

        if (string.Equals(kind, "column", StringComparison.OrdinalIgnoreCase))
        {
            value = ConflictKind.Column;
            return true;
        }

        value = default;
        return false;
    }

    private static string? EmptyToNull(string? value) => string.IsNullOrWhiteSpace(value) ? null : value;

    private static void TryDelete(string dir)
    {
        try
        {
            if (Directory.Exists(dir))
            {
                Directory.Delete(dir, true);
            }
        }
        catch (Exception ex)
        {
            AppLogQuiet(ex);
        }
    }

    private static void AppLogQuiet(Exception ex)
    {
        try
        {
            Console.Error.WriteLine("清理临时目录失败: " + ex.Message);
        }
        catch
        {
            // 临时目录留到下次开机清理即可。
        }
    }
}
