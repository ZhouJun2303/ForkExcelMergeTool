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
        var rows = new List<DiffRow>();

        foreach (var sheet in sheetNames)
        {
            var snapA = sessionA.Sheets.GetValueOrDefault(sheet);
            var snapB = sessionB.Sheets.GetValueOrDefault(sheet);
            var headerA = snapA?.Headers ?? [];
            var headerB = snapB?.Headers ?? [];
            var normA = headerA.Where(h => h.Length > 0).Select(KeyNormalizer.HeaderForCompare).ToHashSet(StringComparer.Ordinal);
            var normB = headerB.Where(h => h.Length > 0).Select(KeyNormalizer.HeaderForCompare).ToHashSet(StringComparer.Ordinal);
            var newCols = headerB.Where(h => h.Length > 0 && !normA.Contains(KeyNormalizer.HeaderForCompare(h))).ToList();
            var delCols = headerA.Where(h => h.Length > 0 && !normB.Contains(KeyNormalizer.HeaderForCompare(h))).ToList();
            if (newCols.Count > 0)
            {
                rows.Add(new DiffRow { Sheet = sheet, Key = "[新增列]", Status = "新增列", Left = "", Right = string.Join(" | ", newCols) });
            }

            if (delCols.Count > 0)
            {
                rows.Add(new DiffRow { Sheet = sheet, Key = "[删除列]", Status = "删除列", Left = string.Join(" | ", delCols), Right = "" });
            }

            var rowsA = snapA is null ? [] : snapA.Rows.Where(r => r.Count > 0 && CellText.From(r[0]).Length > 0).ToList();
            var rowsB = snapB is null ? [] : snapB.Rows.Where(r => r.Count > 0 && CellText.From(r[0]).Length > 0).ToList();
            var dictA = SheetSnapshot.RowsByRawKey(rowsA);
            var dictB = SheetSnapshot.RowsByRawKey(rowsB);
            var maxCol = Math.Max(
                rowsA.Count == 0 ? 1 : rowsA.Max(r => r.Count),
                rowsB.Count == 0 ? 1 : rowsB.Max(r => r.Count));
            foreach (var key in dictA.Keys.Concat(dictB.Keys).Distinct(StringComparer.Ordinal).OrderBy(k => k, StringComparer.Ordinal))
            {
                dictA.TryGetValue(key, out var rowA);
                dictB.TryGetValue(key, out var rowB);
                var valsA = Pad(rowA, maxCol);
                var valsB = Pad(rowB, maxCol);
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

                rows.Add(new DiffRow
                {
                    Sheet = sheet,
                    Key = key,
                    Status = status,
                    Left = string.Join(" | ", valsA),
                    Right = string.Join(" | ", valsB),
                });
            }
        }

        return new CompareResult
        {
            OutputPath = pathOut,
            SheetNames = sheetNames,
            Rows = rows,
            ElapsedMs = (int)(DateTime.UtcNow - started).TotalMilliseconds,
        };
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

        if (names.Count == 0)
        {
            names.Add(a.AllSheetNames.FirstOrDefault() ?? b.AllSheetNames.FirstOrDefault() ?? "Sheet1");
        }

        return names;
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
}
