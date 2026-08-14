using ClosedXML.Excel;
using ExcelMergeFork.Core.Excel;

namespace ExcelMergeFork.Core.Merge;

internal static class WorkbookOps
{
    public static void CopyCellValueAndStyle(IXLCell source, IXLCell dest, XLColor? fontColor = null)
    {
        dest.Value = source.Value;
        dest.Style = source.Style;
        if (fontColor is not null)
        {
            dest.Style.Font.FontColor = fontColor;
        }
    }

    public static void CopyRowWithStyle(IXLWorksheet source, IXLWorksheet dest, int srcRow, int destRow, int maxCol)
    {
        for (var c = 1; c <= maxCol; c++)
        {
            var src = source.Cell(srcRow, c);
            var dst = dest.Cell(destRow, c);
            dst.Value = MergedValue(source, srcRow, c);
            dst.Style = src.Style;
        }

        var height = source.Row(srcRow).Height;
        if (height > 0)
        {
            dest.Row(destRow).Height = height;
        }
    }

    public static void CopyColumnWithStyle(IXLWorksheet source, IXLWorksheet dest, int srcCol, int destCol, int maxRow)
    {
        for (var r = 1; r <= maxRow; r++)
        {
            var src = source.Cell(r, srcCol);
            var dst = dest.Cell(r, destCol);
            dst.Value = MergedValue(source, r, srcCol);
            dst.Style = src.Style;
        }

        dest.Column(destCol).Width = source.Column(srcCol).Width;
    }

    public static XLCellValue MergedValue(IXLWorksheet ws, int row, int col)
    {
        var cell = ws.Cell(row, col);
        if (cell.IsMerged())
        {
            return cell.MergedRange().FirstCell().Value;
        }

        return cell.Value;
    }

    public static void CopyWorksheet(XLWorkbook dest, IXLWorksheet source, string? title = null)
    {
        var name = title ?? source.Name;
        if (dest.Worksheets.Contains(name))
        {
            dest.Worksheets.Delete(name);
        }

        source.CopyTo(dest, name);
    }

    public static void CopyRowMergedRanges(IXLWorksheet source, int srcRow, IXLWorksheet dest, int destRow)
    {
        foreach (var range in source.MergedRanges.ToList())
        {
            var first = range.RangeAddress.FirstAddress;
            var last = range.RangeAddress.LastAddress;
            if (first.RowNumber <= srcRow && srcRow <= last.RowNumber && first.ColumnNumber < last.ColumnNumber)
            {
                try
                {
                    dest.Range(destRow, first.ColumnNumber, destRow, last.ColumnNumber).Merge();
                }
                catch
                {
                    // overlapping merge is ignored
                }
            }
        }
    }

    public static void CopyColMergedRanges(IXLWorksheet source, int srcCol, IXLWorksheet dest, int destCol)
    {
        foreach (var range in source.MergedRanges.ToList())
        {
            var first = range.RangeAddress.FirstAddress;
            var last = range.RangeAddress.LastAddress;
            if (first.ColumnNumber <= srcCol && srcCol <= last.ColumnNumber && first.RowNumber < last.RowNumber)
            {
                try
                {
                    dest.Range(first.RowNumber, destCol, last.RowNumber, destCol).Merge();
                }
                catch
                {
                }
            }
        }
    }

    public static Dictionary<string, int> RowKeyToIndex(IXLWorksheet ws, int maxCol)
    {
        var snapshot = SheetSnapshot.From(ws, maxCol);
        return new Dictionary<string, int>(snapshot.KeyToRowIndex, StringComparer.Ordinal);
    }

    public static void ShiftRowMapAfterDelete(Dictionary<string, int> map, int deletedRow)
    {
        foreach (var key in map.Keys.ToList())
        {
            if (map[key] == deletedRow)
            {
                map.Remove(key);
            }
            else if (map[key] > deletedRow)
            {
                map[key]--;
            }
        }
    }

    public static void ShiftRowMapAfterInsert(Dictionary<string, int> map, int insertAt, int amount = 1)
    {
        foreach (var key in map.Keys.ToList())
        {
            if (map[key] >= insertAt)
            {
                map[key] += amount;
            }
        }
    }

    public static List<(int OutCol, int SrcCol, int BaseCol)> RowCopyPlan(
        IXLWorksheet dest,
        IXLWorksheet source,
        IXLWorksheet? baseline,
        int maxCol)
    {
        var headerOut = SheetSnapshot.From(dest, maxCol).Headers;
        var headerSrc = SheetSnapshot.From(source).Headers;
        var headerBase = baseline is null ? [] : SheetSnapshot.From(baseline).Headers;
        var srcMap = SheetSnapshot.HeaderIndex(headerSrc, true);
        var baseMap = SheetSnapshot.HeaderIndex(headerBase, true);
        var destMap = SheetSnapshot.HeaderIndex(headerOut, true);
        var named = srcMap.Count > 0 && destMap.Count > 0;
        var plan = new List<(int, int, int)>();
        for (var outCol = 1; outCol <= maxCol; outCol++)
        {
            var srcCol = outCol;
            var baseCol = outCol;
            if (named && outCol <= headerOut.Count)
            {
                var norm = KeyNormalizer.HeaderForCompare(headerOut[outCol - 1]);
                if (srcMap.TryGetValue(norm, out var s))
                {
                    srcCol = s;
                }

                if (baseMap.TryGetValue(norm, out var b))
                {
                    baseCol = b;
                }
            }

            plan.Add((outCol, srcCol, baseCol));
        }

        return plan;
    }

    public static void EnsureDirectory(string path)
    {
        var dir = Path.GetDirectoryName(Path.GetFullPath(path));
        if (!string.IsNullOrEmpty(dir))
        {
            Directory.CreateDirectory(dir);
        }
    }

    public static void EnsureSheet(XLWorkbook wb)
    {
        if (!wb.Worksheets.Any())
        {
            wb.AddWorksheet("Data");
        }
    }

    public static XLWorkbook CreateEmpty()
    {
        var wb = new XLWorkbook();
        foreach (var ws in wb.Worksheets.ToList())
        {
            wb.Worksheets.Delete(ws.Name);
        }

        return wb;
    }
}
