using ExcelMergeFork.Core.Compare;

namespace ExcelMergeFork.Api;

public sealed class CompareWorkSession : IDisposable
{
    private bool _disposed;

    private CompareWorkSession(string id, string ownedDir, string leftSource, string rightSource, string leftCopy, string rightCopy)
    {
        Id = id;
        OwnedDir = ownedDir;
        LeftSource = leftSource;
        RightSource = rightSource;
        LeftCopy = leftCopy;
        RightCopy = rightCopy;
        LeftLabel = Path.GetFileName(leftSource);
        RightLabel = Path.GetFileName(rightSource);
        TouchedUtc = DateTime.UtcNow;
    }

    public object Gate { get; } = new();
    public string Id { get; }
    public string OwnedDir { get; }
    public string LeftSource { get; private set; }
    public string RightSource { get; private set; }
    public string LeftCopy { get; private set; }
    public string RightCopy { get; private set; }
    public string LeftLabel { get; private set; }
    public string RightLabel { get; private set; }
    public bool IncludeSame { get; private set; }
    public CompareResult? Result { get; private set; }
    public string? ExportPath { get; private set; }
    public string Status { get; private set; } = "正在计算差异...";
    public DateTime TouchedUtc { get; private set; }

    public static CompareWorkSession Open(string leftSource, string rightSource, bool includeSame)
    {
        var id = Guid.NewGuid().ToString("N");
        var dir = Path.Combine(RepoLayout.WorkRoot, "compare-" + id);
        Directory.CreateDirectory(dir);
        var leftCopy = Path.Combine(dir, CopyName("left", leftSource, rightSource));
        var rightCopy = Path.Combine(dir, CopyName("right", rightSource, leftSource));
        try
        {
            File.Copy(leftSource, leftCopy, true);
            File.Copy(rightSource, rightCopy, true);
            var work = new CompareWorkSession(id, dir, leftSource, rightSource, leftCopy, rightCopy)
            {
                IncludeSame = includeSame,
            };
            work.Compute();
            return work;
        }
        catch
        {
            TryDelete(dir);
            throw;
        }
    }

    public void Swap()
    {
        (LeftSource, RightSource) = (RightSource, LeftSource);
        (LeftCopy, RightCopy) = (RightCopy, LeftCopy);
        (LeftLabel, RightLabel) = (RightLabel, LeftLabel);
        ExportPath = null;
        Compute();
    }

    public void SetIncludeSame(bool includeSame)
    {
        IncludeSame = includeSame;
        ExportPath = null;
        Compute();
    }

    public string Export()
    {
        if (Result is null)
        {
            Compute();
        }

        CompareEngine.WriteExcel(Result!, openFile: false);
        if (Result is null || !File.Exists(Result.OutputPath))
        {
            throw new InvalidOperationException("对比表没有写出来");
        }

        ExportPath = Result.OutputPath;
        Status = "已导出 " + Path.GetFileName(ExportPath);
        return ExportPath;
    }

    public object Snapshot(string? sheet, bool sheetSpecified, string? q, IReadOnlyCollection<string>? statuses, PageQuery page)
    {
        Touch();
        var result = Result ?? throw new InvalidOperationException("对比还没有结果");
        var counts = result.Rows
            .GroupBy(row => row.Status, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        var sheetCounts = result.Rows
            .GroupBy(row => row.Sheet, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);

        string? activeSheet = sheetSpecified ? (string.IsNullOrWhiteSpace(sheet) ? null : sheet) : FirstDiffSheet(result, sheetCounts);
        IEnumerable<(DiffRow Row, int Index)> scoped = result.Rows.Select((row, index) => (row, index));
        if (activeSheet is not null)
        {
            scoped = scoped.Where(item => string.Equals(item.Row.Sheet, activeSheet, StringComparison.Ordinal));
        }

        var scopedCount = scoped.Count();
        IEnumerable<(DiffRow Row, int Index)> rows = scoped;
        if (statuses is not null)
        {
            var set = new HashSet<string>(statuses, StringComparer.Ordinal);
            rows = rows.Where(item => set.Contains(item.Row.Status));
        }

        if (!string.IsNullOrWhiteSpace(q))
        {
            rows = rows.Where(item =>
                item.Row.Key.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                item.Row.Sheet.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                item.Row.Left.Contains(q, StringComparison.OrdinalIgnoreCase) ||
                item.Row.Right.Contains(q, StringComparison.OrdinalIgnoreCase));
        }

        var filtered = rows.ToList();
        var slice = filtered.Skip(page.Offset).Take(page.Limit).Select(item =>
        {
            var changed = item.Row.Cells.Where(cell => cell.Changed).ToList();
            return new
            {
                item.Index,
                item.Row.Sheet,
                item.Row.Key,
                item.Row.Status,
                Left = Brief(item.Row.Left),
                Right = Brief(item.Row.Right),
                item.Row.LeftRow,
                item.Row.RightRow,
                Location = FormatLocation(item.Row.LeftRow, item.Row.RightRow),
                Preview = PreviewOf(changed),
            };
        }).ToList();

        var empty = EmptyState(result, scopedCount, filtered.Count);
        var scopeName = activeSheet ?? "全部工作表";
        var statusText = $"{scopeName}：{scopedCount} 处差异，当前显示 {filtered.Count} 处。全部 {result.Rows.Count} 处，耗时 {result.ElapsedMs} ms";
        if (result.Rows.Count == 0)
        {
            statusText = empty.Title;
        }

        var sheets = new List<object>
        {
            new
            {
                Name = "全部工作表",
                Sheet = (string?)null,
                Count = result.Rows.Count,
                Badge = result.Rows.Count == 0 ? "无差异" : result.Rows.Count + " 处",
                Current = activeSheet is null,
            },
        };
        foreach (var name in result.SheetNames)
        {
            var count = sheetCounts.GetValueOrDefault(name);
            sheets.Add(new
            {
                Name = name,
                Sheet = (string?)name,
                Count = count,
                Badge = count == 0 ? "无差异" : count + " 处",
                Current = string.Equals(name, activeSheet, StringComparison.Ordinal),
            });
        }

        return new
        {
            Id,
            Status = statusText,
            ElapsedMs = result.ElapsedMs,
            LeftPath = LeftSource,
            RightPath = RightSource,
            LocalPath = RightSource,
            RemotePath = LeftSource,
            LocalLabel = RightLabel,
            RemoteLabel = LeftLabel,
            LeftLabel,
            RightLabel,
            IncludeSame,
            ActiveSheet = activeSheet,
            result.SheetNames,
            Skipped = result.SkippedSheetNames,
            SkippedHint = result.SkippedSheetNames.Count == 0
                ? ""
                : "已跳过以 # 开头的工作表：" + string.Join("、", result.SkippedSheetNames),
            Counts = counts,
            Sheets = sheets,
            Scoped = scopedCount,
            ShowEmpty = empty.Show,
            EmptyTitle = empty.Title,
            EmptyMessage = empty.Message,
            ItemTotal = filtered.Count,
            Total = result.Rows.Count,
            page.Offset,
            page.Limit,
            Rows = slice,
            ExportPath,
            DownloadUrl = "/api/compare/sessions/" + Id + "/download",
        };
    }

    public object Detail(int index, bool changedOnly)
    {
        Touch();
        var result = Result ?? throw new InvalidOperationException("对比还没有结果");
        if (index < 0 || index >= result.Rows.Count)
        {
            throw new InvalidOperationException("差异行不存在");
        }

        var row = result.Rows[index];
        var cells = changedOnly ? row.Cells.Where(cell => cell.Changed) : row.Cells;
        return new
        {
            Index = index,
            row.Sheet,
            row.Key,
            row.Status,
            row.Left,
            row.Right,
            row.LeftRow,
            row.RightRow,
            ChangedOnly = changedOnly,
            CellTotal = row.Cells.Count,
            Changed = row.Cells.Count(cell => cell.Changed),
            Cells = cells.Take(200).Select(cell => new
            {
                cell.Header,
                Left = cell.Left.Length <= 2000 ? cell.Left : cell.Left[..2000] + "…",
                Right = cell.Right.Length <= 2000 ? cell.Right : cell.Right[..2000] + "…",
                cell.Changed,
            }).ToList(),
        };
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
            Result = null;
        }

        TryDelete(OwnedDir);
    }

    private void Compute()
    {
        Result = CompareEngine.Compute(LeftCopy, RightCopy, IncludeSame);
        Status = "对比完成，" + Result.Rows.Count + " 条差异";
    }

    private static string CopyName(string role, string path, string other)
    {
        var name = RepoLayout.SafeFileName(Path.GetFileName(path));
        var otherName = RepoLayout.SafeFileName(Path.GetFileName(other));
        if (!name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) &&
            !name.EndsWith(".xltx", StringComparison.OrdinalIgnoreCase))
        {
            name += ".xlsx";
        }

        return string.Equals(name, otherName, StringComparison.OrdinalIgnoreCase) ? role + "_" + name : name;
    }

    private static string? FirstDiffSheet(CompareResult result, Dictionary<string, int> sheetCounts)
    {
        foreach (var name in result.SheetNames)
        {
            if (sheetCounts.GetValueOrDefault(name) > 0)
            {
                return name;
            }
        }

        return null;
    }

    private static (bool Show, string Title, string Message) EmptyState(CompareResult result, int scoped, int shown)
    {
        if (result.Rows.Count == 0)
        {
            var skippedOnly = result.SkippedSheetNames.Count > 0 && result.SheetNames.Count == 0;
            return (true, "没有差异", skippedOnly
                ? "这些工作表以 # 开头，对比时会跳过：" + string.Join("、", result.SkippedSheetNames)
                : "按第一列主键逐条对比后，两边内容一致。");
        }

        if (shown == 0)
        {
            return scoped == 0
                ? (true, "这个工作表没有差异", "左侧可以切换到其他工作表。有差异的表会标出数量。")
                : (true, "没有符合筛选的差异", "试试清空搜索，或勾选上方的差异类型。");
        }

        return (false, "", "");
    }

    private static string FormatLocation(int remoteRow, int localRow)
    {
        if (localRow > 0 && remoteRow > 0 && localRow != remoteRow)
        {
            return "本地 " + localRow + " 行 · 线上 " + remoteRow + " 行";
        }

        if (localRow > 0 && remoteRow > 0)
        {
            return "第 " + localRow + " 行";
        }

        if (localRow > 0)
        {
            return "本地第 " + localRow + " 行";
        }

        if (remoteRow > 0)
        {
            return "线上第 " + remoteRow + " 行";
        }

        return "";
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

    private static string Brief(string value) => value.Length <= 180 ? value : value[..180] + "…";

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
            Console.Error.WriteLine("清理临时目录失败: " + ex.Message);
        }
    }
}
