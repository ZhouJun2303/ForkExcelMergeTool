using System.Globalization;
using ClosedXML.Excel;
using ExcelMergeFork.Core.Excel;
using ExcelMergeFork.Core.Merge;

namespace ExcelMergeFork.Core.Compare;

public static class CompareEngine
{
    public static CompareResult Compute(string pathA, string pathB, bool includeSame = false)
    {
        var started = DateTime.UtcNow;
        var outDir = Path.GetDirectoryName(Path.GetFullPath(pathA)) ?? ".";
        var baseName = Path.GetFileNameWithoutExtension(pathA);
        var pathOut = Path.Combine(outDir, baseName + AppConstants.CompareSuffix + ".xlsx");

        using var sessionA = WorkbookSession.Open(pathA);
        using var sessionB = WorkbookSession.Open(pathB);
        var sheetNames = UnionSheets(sessionA, sessionB);
        var skipped = sessionA.AllSheetNames
            .Concat(sessionB.AllSheetNames)
            .Where(SheetFilter.ShouldSkip)
            .Distinct(StringComparer.Ordinal)
            .ToList();
        var rows = new List<DiffRow>();

        foreach (var sheet in sheetNames)
        {
            CompareSheet(sheet, sessionA.Sheets.GetValueOrDefault(sheet), sessionB.Sheets.GetValueOrDefault(sheet), rows, includeSame);
        }

        return new CompareResult
        {
            OutputPath = pathOut,
            SheetNames = sheetNames,
            SkippedSheetNames = skipped,
            Rows = rows,
            ElapsedMs = (int)(DateTime.UtcNow - started).TotalMilliseconds,
        };
    }

    private static void CompareSheet(string sheet, SheetSnapshot? snapA, SheetSnapshot? snapB, List<DiffRow> rows, bool includeSame)
    {
        var headerA = snapA?.Headers ?? [];
        var headerB = snapB?.Headers ?? [];
        var normA = headerA.Where(h => h.Length > 0).Select(KeyNormalizer.HeaderForCompare).ToHashSet(StringComparer.Ordinal);
        var normB = headerB.Where(h => h.Length > 0).Select(KeyNormalizer.HeaderForCompare).ToHashSet(StringComparer.Ordinal);
        var newCols = headerB.Where(h => h.Length > 0 && !normA.Contains(KeyNormalizer.HeaderForCompare(h))).ToList();
        var delCols = headerA.Where(h => h.Length > 0 && !normB.Contains(KeyNormalizer.HeaderForCompare(h))).ToList();
        if (newCols.Count > 0)
        {
            rows.Add(new DiffRow
            {
                Sheet = sheet,
                Key = "[新增列]",
                Status = "新增列",
                Left = "",
                Right = string.Join(" | ", newCols),
                Cells = newCols.Select(name => new DiffCell { Header = name, Left = "", Right = name, Changed = true }).ToList(),
            });
        }

        if (delCols.Count > 0)
        {
            rows.Add(new DiffRow
            {
                Sheet = sheet,
                Key = "[删除列]",
                Status = "删除列",
                Left = string.Join(" | ", delCols),
                Right = "",
                Cells = delCols.Select(name => new DiffCell { Header = name, Left = name, Right = "", Changed = true }).ToList(),
            });
        }

        var widthA = Width(snapA);
        var widthB = Width(snapB);
        var maxCol = Math.Max(widthA, widthB);
        // 同一主键出现多次时按出现次序逐条对齐。只留最后一行会把前面的修改和多出来的重复行藏掉。
        var sideA = Enumerate(snapA, maxCol);
        var sideB = Enumerate(snapB, maxCol);
        var mapA = Map(sideA);
        var mapB = Map(sideB);
        foreach (var id in Order(sideA, sideB))
        {
            mapA.TryGetValue(id, out var rowA);
            mapB.TryGetValue(id, out var rowB);
            var valsA = rowA?.Values ?? Pad(null, maxCol);
            var valsB = rowB?.Values ?? Pad(null, maxCol);
            string status;
            if (rowA is null)
            {
                status = "新增行";
            }
            else if (rowB is null)
            {
                status = "删除行";
            }
            else if (!valsA.SequenceEqual(valsB))
            {
                status = "修改";
            }
            else
            {
                status = "相同";
            }

            if (status == "相同" && !includeSame)
            {
                continue;
            }

            var ordinal = rowA?.Ordinal ?? rowB?.Ordinal ?? 1;
            var key = rowA?.Key ?? rowB?.Key ?? "";
            rows.Add(new DiffRow
            {
                Sheet = sheet,
                Key = ordinal <= 1 ? key : key + " #" + ordinal.ToString(CultureInfo.InvariantCulture),
                Status = status,
                Left = string.Join(" | ", valsA),
                Right = string.Join(" | ", valsB),
                LeftRow = rowA?.ExcelRow ?? 0,
                RightRow = rowB?.ExcelRow ?? 0,
                Cells = BuildCells(valsA, valsB, headerA, headerB),
            });
        }
    }

    public static void WriteExcel(CompareResult result, bool openFile = false)
    {
        using var wb = WorkbookOps.CreateEmpty();
        var sheets = new Dictionary<string, IXLWorksheet>(StringComparer.Ordinal);
        foreach (var name in result.SheetNames)
        {
            var ws = wb.AddWorksheet(SafeSheetName(name));
            ws.Cell(1, 1).Value = "[Key]";
            ws.Cell(1, 2).Value = "[A-LEFT]";
            ws.Cell(1, 3).Value = "[B-RIGHT]";
            ws.Cell(1, 4).Value = "[Status]";
            sheets[name] = ws;
        }

        var used = sheets.ToDictionary(kv => kv.Key, kv => 1, StringComparer.Ordinal);
        foreach (var row in result.Rows)
        {
            if (!sheets.TryGetValue(row.Sheet, out var ws))
            {
                continue;
            }

            var r = used[row.Sheet] + 1;
            used[row.Sheet] = r;
            ws.Cell(r, 1).Value = row.Key;
            ws.Cell(r, 2).Value = row.Left;
            ws.Cell(r, 3).Value = row.Right;
            ws.Cell(r, 4).Value = row.Status;
            var color = row.Status is "新增行" or "新增列"
                ? AppConstants.ColorNew
                : row.Status is "删除行" or "删除列"
                    ? AppConstants.ColorConflict
                    : row.Status == "修改"
                        ? AppConstants.ColorChanged
                        : null;
            if (color is not null)
            {
                ws.Range(r, 1, r, 4).Style.Fill.BackgroundColor = XLColor.FromHtml("#" + color);
            }
        }

        if (wb.Worksheets.Count == 0)
        {
            wb.AddWorksheet("Sheet1");
        }

        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(result.OutputPath)) ?? ".");
        wb.SaveAs(result.OutputPath);
        if (openFile)
        {
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(result.OutputPath) { UseShellExecute = true });
            }
            catch
            {
            }
        }
    }

    public static int CompareAndWrite(string pathA, string pathB, bool openFile = false, bool includeSame = false)
    {
        if (!File.Exists(pathA) || !File.Exists(pathB))
        {
            AppLog.Error("对比失败: 文件不存在");
            return 1;
        }

        try
        {
            var result = Compute(pathA, pathB, includeSame);
            WriteExcel(result, openFile);
            AppLog.Info("对比模式 输出: " + result.OutputPath);
            return 0;
        }
        catch (Exception ex)
        {
            AppLog.Exception("对比失败", ex);
            return 2;
        }
    }

    private static List<string> UnionSheets(WorkbookSession a, WorkbookSession b)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var names = new List<string>();
        foreach (var name in a.SheetNames.Concat(b.SheetNames))
        {
            if (seen.Add(name))
            {
                names.Add(name);
            }
        }

        if (names.Count == 0 && a.AllSheetNames.Count == 0 && b.AllSheetNames.Count == 0)
        {
            names.Add("Sheet1");
        }

        return names;
    }

    private static int Width(SheetSnapshot? snap)
    {
        if (snap is null || snap.Rows.Count == 0)
        {
            return 1;
        }

        var width = 1;
        foreach (var row in snap.Rows)
        {
            if (row.Count > 0 && CellText.From(row[0]).Length > 0 && row.Count > width)
            {
                width = row.Count;
            }
        }

        return width;
    }

    private static List<OccRow> Enumerate(SheetSnapshot? snap, int maxCol)
    {
        var list = new List<OccRow>();
        if (snap is null)
        {
            return list;
        }

        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        for (var i = 0; i < snap.Rows.Count; i++)
        {
            var raw = snap.Rows[i];
            var key = raw.Count > 0 ? CellText.From(raw[0]) : "";
            if (key.Length == 0)
            {
                continue;
            }

            var ordinal = counts.GetValueOrDefault(key) + 1;
            counts[key] = ordinal;
            list.Add(new OccRow
            {
                Key = key,
                Ordinal = ordinal,
                ExcelRow = i < snap.RowIndices.Count ? snap.RowIndices[i] : i + 1,
                Values = Pad(raw, maxCol),
            });
        }

        return list;
    }

    private static Dictionary<string, OccRow> Map(List<OccRow> rows)
    {
        var dict = new Dictionary<string, OccRow>(StringComparer.Ordinal);
        foreach (var row in rows)
        {
            dict[Id(row)] = row;
        }

        return dict;
    }

    private static List<string> Order(List<OccRow> sideA, List<OccRow> sideB)
    {
        var order = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var row in sideA.Concat(sideB))
        {
            var id = Id(row);
            if (seen.Add(id))
            {
                order.Add(id);
            }
        }

        return order;
    }

    private static string Id(OccRow row) =>
        row.Ordinal.ToString(CultureInfo.InvariantCulture) + "\u001f" + row.Key;

    private static List<DiffCell> BuildCells(
        IReadOnlyList<string> left,
        IReadOnlyList<string> right,
        IReadOnlyList<string> headerA,
        IReadOnlyList<string> headerB)
    {
        var n = Math.Max(left.Count, right.Count);
        var cells = new List<DiffCell>(n);
        for (var i = 0; i < n; i++)
        {
            var a = i < left.Count ? left[i] : "";
            var b = i < right.Count ? right[i] : "";
            cells.Add(new DiffCell
            {
                Header = HeaderAt(headerA, headerB, i),
                Left = a,
                Right = b,
                Changed = a != b,
            });
        }

        return cells;
    }

    private static string HeaderAt(IReadOnlyList<string> headerA, IReadOnlyList<string> headerB, int index)
    {
        var a = index < headerA.Count ? headerA[index].Trim() : "";
        var b = index < headerB.Count ? headerB[index].Trim() : "";
        if (a.Length > 0 && b.Length > 0 && !string.Equals(a, b, StringComparison.Ordinal))
        {
            return a + " / " + b;
        }

        if (a.Length > 0)
        {
            return a;
        }

        if (b.Length > 0)
        {
            return b;
        }

        return "列" + (index + 1).ToString(CultureInfo.InvariantCulture);
    }

    private static List<string> Pad(List<object?>? row, int maxCol)
    {
        var vals = new List<string>(maxCol);
        for (var i = 0; i < maxCol; i++)
        {
            vals.Add(row is not null && i < row.Count ? CellText.From(row[i]) : "");
        }

        return vals;
    }

    private static string SafeSheetName(string name)
    {
        var trimmed = name.Length <= 31 ? name : name[..31];
        return string.IsNullOrWhiteSpace(trimmed) ? "Sheet1" : trimmed;
    }

    private sealed class OccRow
    {
        public required string Key { get; init; }
        public required int Ordinal { get; init; }
        public required int ExcelRow { get; init; }
        public required List<string> Values { get; init; }
    }
}
