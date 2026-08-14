using ClosedXML.Excel;
using ExcelMergeFork.Core.Excel;

namespace ExcelMergeFork.Core.Merge;

public static class MergePipeline
{
    private static readonly XLColor NewFont = XLColor.FromHtml("#" + AppConstants.FontNew);
    private static readonly XLColor ModifiedFont = XLColor.FromHtml("#" + AppConstants.FontModified);

    public static void RunByMode(
        string pathLocal,
        string pathBase,
        string pathRemote,
        string pathMerged,
        string mode,
        string baseSide,
        IReadOnlyList<MergeChoice>? choices)
    {
        var pathBaseSide = baseSide == "remote" ? pathRemote : pathLocal;
        var pathOtherSide = baseSide == "remote" ? pathLocal : pathRemote;
        switch (mode.ToUpperInvariant())
        {
            case "A":
                MergeNewRowsClassic(pathBaseSide, pathOtherSide, pathMerged);
                break;
            case "B":
                MergeNewColumnsClassic(pathBaseSide, pathOtherSide, pathMerged);
                break;
            case "C":
                MergeNewSheetsClassic(pathBaseSide, pathOtherSide, pathMerged);
                break;
            case "D":
                using (var session = new MergeSession(pathLocal, pathBase, pathRemote))
                {
                    var auto = ConflictDetector.DetectAutoActions(session);
                    ApplyChoicesClassic(pathLocal, pathRemote, pathMerged, pathBase, ConflictDetector.MergeChoices(choices, auto));
                }
                break;
            default:
                RunByOptions(pathLocal, pathBase, pathRemote, pathMerged, MergeOptions.FromLetters(["E", "G"], baseSide), choices);
                break;
        }
    }

    public static void RunByOptions(
        string pathLocal,
        string pathBase,
        string pathRemote,
        string pathMerged,
        MergeOptions options,
        IReadOnlyList<MergeChoice>? choices)
    {
        var pathBaseSide = options.BaseSide == "remote" ? pathRemote : pathLocal;
        var pathOtherSide = options.BaseSide == "remote" ? pathLocal : pathRemote;
        var temp = Path.Combine(Path.GetTempPath(), "merge_opt_" + Guid.NewGuid().ToString("N") + ".xlsx");
        File.Copy(pathBaseSide, temp, overwrite: true);
        try
        {
            using var wbOut = new XLWorkbook(temp);
            using var wbLocal = new XLWorkbook(pathLocal);
            using var wbBase = new XLWorkbook(pathBase);
            using var wbRemote = new XLWorkbook(pathRemote);
            var wbOther = options.BaseSide == "remote" ? wbLocal : wbRemote;

            if (!options.SkipNewRows)
            {
                InsertNewRows(wbOut, wbOther);
            }

            if (options.DeleteMissingRows)
            {
                DeleteMissingRows(wbOut, wbOther);
            }

            if (!options.SkipNewColumns)
            {
                InsertNewColumns(wbOut, wbOther);
            }

            if (options.DeleteMissingColumns)
            {
                DeleteMissingColumns(wbOut, wbOther);
            }

            if (options.AddNewSheets)
            {
                InsertNewSheets(wbOut, wbOther);
            }

            if (options.DeleteMissingSheets)
            {
                DeleteMissingSheets(wbOut, wbOther);
            }

            if (options.ResolveConflicts)
            {
                using var session = new MergeSession(pathLocal, pathBase, pathRemote);
                var auto = ConflictDetector.DetectAutoActions(session);
                ApplyChoices(wbOut, wbLocal, wbBase, wbRemote, ConflictDetector.MergeChoices(choices, auto));
            }

            WorkbookOps.EnsureSheet(wbOut);
            WorkbookOps.EnsureDirectory(pathMerged);
            wbOut.SaveAs(pathMerged);
        }
        finally
        {
            try { File.Delete(temp); } catch { /* ignore */ }
        }
    }

    internal static int InsertNewRows(XLWorkbook wbIn, XLWorkbook wbOther)
    {
        var inserted = 0;
        foreach (var wsIn in wbIn.Worksheets.ToList())
        {
            if (!wbOther.Worksheets.Contains(wsIn.Name) || SheetFilter.ShouldSkip(wsIn.Name))
            {
                continue;
            }

            var wsO = wbOther.Worksheet(wsIn.Name);
            var maxCol = Math.Max(wsIn.LastColumnUsed()?.ColumnNumber() ?? 1, wsO.LastColumnUsed()?.ColumnNumber() ?? 1);
            var rowsIn = SheetSnapshot.From(wsIn, maxCol);
            var rowsO = SheetSnapshot.From(wsO, maxCol);
            var baseKeys = rowsIn.OrderedKeysNormalized.ToHashSet(StringComparer.Ordinal);
            var keyToRowIn = new Dictionary<string, int>(StringComparer.Ordinal);
            for (var i = 0; i < rowsIn.Rows.Count; i++)
            {
                var key = rowsIn.Rows[i].Count > 0 ? KeyNormalizer.Normalize(rowsIn.Rows[i][0]) : "";
                if (key.Length > 0)
                {
                    keyToRowIn[key] = rowsIn.RowIndices[i];
                }
            }

            var newKeys = rowsO.OrderedKeysNormalized.Where(k => !baseKeys.Contains(k)).ToList();
            var mergedOrdered = KeyNormalizer.MergeOrdered(rowsIn.OrderedKeysNormalized, newKeys);
            var keyToRowsOther = new Dictionary<string, List<int>>(StringComparer.Ordinal);
            for (var i = 0; i < rowsO.Rows.Count; i++)
            {
                var key = rowsO.Rows[i].Count > 0 ? KeyNormalizer.Normalize(rowsO.Rows[i][0]) : "";
                if (key.Length > 0)
                {
                    if (!keyToRowsOther.TryGetValue(key, out var list))
                    {
                        list = [];
                        keyToRowsOther[key] = list;
                    }

                    list.Add(rowsO.RowIndices[i]);
                }
            }

            int? lastBaseRow = null;
            var inserts = new List<(int After, string Key, int Order)>();
            for (var order = 0; order < mergedOrdered.Count; order++)
            {
                var key = mergedOrdered[order];
                if (keyToRowIn.TryGetValue(key, out var rowIn))
                {
                    lastBaseRow = rowIn;
                }
                else if (keyToRowsOther.ContainsKey(key))
                {
                    inserts.Add((lastBaseRow ?? 0, key, order));
                }
            }

            foreach (var (after, key, _) in inserts
                         .OrderByDescending(x => x.After >= 1 ? x.After : int.MinValue)
                         .ThenByDescending(x => x.Order))
            {
                var otherRows = keyToRowsOther[key];
                var insertAt = after >= 1 ? after + 1 : 2;
                var count = otherRows.Count;
                if (count <= 0)
                {
                    continue;
                }

                wsIn.Row(insertAt).InsertRowsAbove(count);
                for (var i = 0; i < count; i++)
                {
                    var srcRow = otherRows[count - 1 - i];
                    var destRow = insertAt + (count - 1 - i);
                    WorkbookOps.CopyRowWithStyle(wsO, wsIn, srcRow, destRow, maxCol);
                    WorkbookOps.CopyRowMergedRanges(wsO, srcRow, wsIn, destRow);
                }

                inserted += count;
            }
        }

        return inserted;
    }

    internal static int InsertNewColumns(XLWorkbook wbIn, XLWorkbook wbOther)
    {
        var inserted = 0;
        foreach (var wsIn in wbIn.Worksheets.ToList())
        {
            if (!wbOther.Worksheets.Contains(wsIn.Name) || SheetFilter.ShouldSkip(wsIn.Name))
            {
                continue;
            }

            var wsO = wbOther.Worksheet(wsIn.Name);
            var maxCol = Math.Max(wsIn.LastColumnUsed()?.ColumnNumber() ?? 1, wsO.LastColumnUsed()?.ColumnNumber() ?? 1);
            var headerIn = SheetSnapshot.From(wsIn, maxCol).Headers;
            var headerO = SheetSnapshot.From(wsO, maxCol).Headers;
            var headerInMap = SheetSnapshot.HeaderIndex(headerIn, true);
            var otherOrdered = headerO.Where(h => h.Length > 0).ToList();
            var newCols = otherOrdered.Where(h => !headerInMap.ContainsKey(KeyNormalizer.HeaderForCompare(h))).ToList();
            var merged = KeyNormalizer.MergeOrdered(headerIn.Where(h => h.Length > 0), newCols);
            var colToIdxO = new Dictionary<string, int>(StringComparer.Ordinal);
            for (var i = 0; i < headerO.Count; i++)
            {
                if (headerO[i].Length > 0)
                {
                    colToIdxO[headerO[i]] = i + 1;
                }
            }

            int? lastBaseCol = null;
            var inserts = new List<(int After, string Key)>();
            foreach (var header in merged)
            {
                if (headerInMap.TryGetValue(KeyNormalizer.HeaderForCompare(header), out var idxIn))
                {
                    lastBaseCol = idxIn;
                }
                else if (colToIdxO.ContainsKey(header))
                {
                    inserts.Add((lastBaseCol ?? 0, header));
                }
            }

            var maxRow = Math.Max(wsIn.LastRowUsed()?.RowNumber() ?? 1, wsO.LastRowUsed()?.RowNumber() ?? 1);
            foreach (var (after, colKey) in inserts.OrderByDescending(x => x.After))
            {
                var colO = colToIdxO[colKey];
                var insertAt = after >= 1 ? after + 1 : (wsIn.LastColumnUsed()?.ColumnNumber() ?? 1) + 1;
                wsIn.Column(insertAt).InsertColumnsBefore(1);
                WorkbookOps.CopyColumnWithStyle(wsO, wsIn, colO, insertAt, maxRow);
                WorkbookOps.CopyColMergedRanges(wsO, colO, wsIn, insertAt);
                inserted++;
            }
        }

        return inserted;
    }

    internal static int DeleteMissingRows(XLWorkbook wbIn, XLWorkbook wbOther)
    {
        var deleted = 0;
        foreach (var wsIn in wbIn.Worksheets.ToList())
        {
            if (!wbOther.Worksheets.Contains(wsIn.Name))
            {
                continue;
            }

            var wsO = wbOther.Worksheet(wsIn.Name);
            var maxCol = Math.Max(wsIn.LastColumnUsed()?.ColumnNumber() ?? 1, wsO.LastColumnUsed()?.ColumnNumber() ?? 1);
            var otherKeys = SheetSnapshot.From(wsO, maxCol).OrderedKeysNormalized.ToHashSet(StringComparer.Ordinal);
            var rowsIn = SheetSnapshot.From(wsIn, maxCol);
            var toDelete = new HashSet<int>();
            for (var i = 0; i < rowsIn.Rows.Count; i++)
            {
                var key = rowsIn.Rows[i].Count > 0 ? KeyNormalizer.Normalize(rowsIn.Rows[i][0]) : "";
                if (key.Length > 0 && !otherKeys.Contains(key))
                {
                    toDelete.Add(rowsIn.RowIndices[i]);
                }
            }

            foreach (var row in toDelete.OrderByDescending(x => x))
            {
                wsIn.Row(row).Delete();
                deleted++;
            }
        }

        return deleted;
    }

    internal static int DeleteMissingColumns(XLWorkbook wbIn, XLWorkbook wbOther)
    {
        var deleted = 0;
        foreach (var wsIn in wbIn.Worksheets.ToList())
        {
            if (!wbOther.Worksheets.Contains(wsIn.Name))
            {
                continue;
            }

            var wsO = wbOther.Worksheet(wsIn.Name);
            var maxCol = Math.Max(wsIn.LastColumnUsed()?.ColumnNumber() ?? 1, wsO.LastColumnUsed()?.ColumnNumber() ?? 1);
            var headerIn = SheetSnapshot.From(wsIn, maxCol).Headers;
            var otherNorm = SheetSnapshot.From(wsO, maxCol).Headers
                .Where(h => h.Length > 0)
                .Select(KeyNormalizer.HeaderForCompare)
                .ToHashSet(StringComparer.Ordinal);
            var toDelete = new List<int>();
            for (var c = headerIn.Count - 1; c >= 0; c--)
            {
                var header = headerIn[c];
                if (header.Length > 0 && !otherNorm.Contains(KeyNormalizer.HeaderForCompare(header)) && c + 1 != 1)
                {
                    toDelete.Add(c + 1);
                }
            }

            foreach (var col in toDelete.OrderByDescending(x => x))
            {
                wsIn.Column(col).Delete();
                deleted++;
            }
        }

        return deleted;
    }

    internal static List<string> InsertNewSheets(XLWorkbook wbIn, XLWorkbook wbOther)
    {
        var baseSheets = wbIn.Worksheets.Select(ws => ws.Name).Where(n => !SheetFilter.ShouldSkip(n)).ToHashSet(StringComparer.Ordinal);
        var added = new List<string>();
        foreach (var ws in wbOther.Worksheets)
        {
            if (SheetFilter.ShouldSkip(ws.Name) || baseSheets.Contains(ws.Name))
            {
                continue;
            }

            WorkbookOps.CopyWorksheet(wbIn, ws, ws.Name);
            added.Add(ws.Name);
        }

        return added;
    }

    internal static List<string> DeleteMissingSheets(XLWorkbook wbIn, XLWorkbook wbOther)
    {
        var other = wbOther.Worksheets.Select(ws => ws.Name).Where(n => !SheetFilter.ShouldSkip(n)).ToHashSet(StringComparer.Ordinal);
        var deleted = new List<string>();
        foreach (var ws in wbIn.Worksheets.ToList())
        {
            if (!other.Contains(ws.Name))
            {
                deleted.Add(ws.Name);
                wbIn.Worksheets.Delete(ws.Name);
            }
        }

        WorkbookOps.EnsureSheet(wbIn);
        return deleted;
    }

    internal static int ApplyChoices(
        XLWorkbook wbOut,
        XLWorkbook wbLocal,
        XLWorkbook wbBase,
        XLWorkbook wbRemote,
        IReadOnlyList<MergeChoice> choices)
    {
        if (choices.Count == 0)
        {
            return 0;
        }

        var cache = new Dictionary<string, SheetChoiceCache?>(StringComparer.Ordinal);
        var applied = 0;
        foreach (var item in choices)
        {
            if (!cache.TryGetValue(item.Sheet, out var sheet))
            {
                sheet = BuildCache(wbOut, wbLocal, wbBase, wbRemote, item.Sheet);
                cache[item.Sheet] = sheet;
            }

            if (sheet is null)
            {
                continue;
            }

            if (item.Kind == ConflictKind.Column)
            {
                if (ApplyColumnChoice(sheet, item))
                {
                    applied++;
                }

                continue;
            }

            if (ApplyRowChoice(sheet, item))
            {
                applied++;
            }
        }

        return applied;
    }

    private static bool ApplyRowChoice(SheetChoiceCache cache, MergeChoice item)
    {
        var keyNorm = KeyNormalizer.Normalize(item.Key);
        if (keyNorm.Length == 0)
        {
            keyNorm = item.Key;
        }

        cache.RowOut.TryGetValue(keyNorm, out var rowOut);
        var source = item.Choice == "remote" ? cache.Remote : cache.Local;
        var sourceRows = item.Choice == "remote" ? cache.RowRemote : cache.RowLocal;
        sourceRows.TryGetValue(keyNorm, out var rowSrc);
        if (source is null || rowSrc == 0)
        {
            if (rowOut != 0)
            {
                cache.Out.Row(rowOut).Delete();
                WorkbookOps.ShiftRowMapAfterDelete(cache.RowOut, rowOut);
                return true;
            }

            return false;
        }

        if (rowOut == 0)
        {
            rowOut = Math.Min(rowSrc, (cache.Out.LastRowUsed()?.RowNumber() ?? 0) + 1);
            cache.Out.Row(rowOut).InsertRowsAbove(1);
            WorkbookOps.ShiftRowMapAfterInsert(cache.RowOut, rowOut);
            cache.RowOut[keyNorm] = rowOut;
        }

        var maxCol = cache.Out.LastColumnUsed()?.ColumnNumber() ?? 1;
        var plan = WorkbookOps.RowCopyPlan(cache.Out, source, cache.Base, maxCol);
        if (item.AutoType is "take_local" or "take_remote")
        {
            cache.RowBase.TryGetValue(keyNorm, out var rowBase);
            foreach (var (outCol, srcCol, baseCol) in plan)
            {
                var srcCell = source.Cell(rowSrc, srcCol);
                var baseVal = cache.Base is not null && rowBase != 0
                    ? CellText.From(CellText.ToObject(cache.Base.Cell(rowBase, baseCol).Value))
                    : "";
                if (CellText.From(CellText.ToObject(srcCell.Value)) == baseVal)
                {
                    continue;
                }

                WorkbookOps.CopyCellValueAndStyle(srcCell, cache.Out.Cell(rowOut, outCol), ModifiedFont);
            }
        }
        else
        {
            foreach (var (outCol, srcCol, _) in plan)
            {
                WorkbookOps.CopyCellValueAndStyle(source.Cell(rowSrc, srcCol), cache.Out.Cell(rowOut, outCol), ModifiedFont);
            }
        }

        return true;
    }

    private static bool ApplyColumnChoice(SheetChoiceCache cache, MergeChoice item)
    {
        EnsureHeaders(cache);
        var norm = KeyNormalizer.HeaderForCompare(item.Key);
        cache.HeaderOut.TryGetValue(norm, out var colOut);
        var source = item.Choice == "remote" ? cache.Remote : cache.Local;
        var sourceMap = item.Choice == "remote" ? cache.HeaderRemote : cache.HeaderLocal;
        sourceMap.TryGetValue(norm, out var colSrc);
        if (source is null || colSrc == 0)
        {
            return false;
        }

        if (colOut == 0)
        {
            colOut = (cache.Out.LastColumnUsed()?.ColumnNumber() ?? 0) + 1;
            cache.HeaderOut[norm] = colOut;
        }

        var maxRow = cache.Out.LastRowUsed()?.RowNumber() ?? 1;
        for (var r = 1; r <= maxRow; r++)
        {
            WorkbookOps.CopyCellValueAndStyle(source.Cell(r, colSrc), cache.Out.Cell(r, colOut), ModifiedFont);
        }

        return true;
    }

    private static SheetChoiceCache? BuildCache(
        XLWorkbook wbOut,
        XLWorkbook wbLocal,
        XLWorkbook wbBase,
        XLWorkbook wbRemote,
        string sheet)
    {
        if (!wbOut.Worksheets.Contains(sheet))
        {
            return null;
        }

        var wsOut = wbOut.Worksheet(sheet);
        var maxCol = wsOut.LastColumnUsed()?.ColumnNumber() ?? 1;
        return new SheetChoiceCache
        {
            Out = wsOut,
            Local = wbLocal.Worksheets.Contains(sheet) ? wbLocal.Worksheet(sheet) : null,
            Base = wbBase.Worksheets.Contains(sheet) ? wbBase.Worksheet(sheet) : null,
            Remote = wbRemote.Worksheets.Contains(sheet) ? wbRemote.Worksheet(sheet) : null,
            RowOut = WorkbookOps.RowKeyToIndex(wsOut, maxCol),
            RowLocal = wbLocal.Worksheets.Contains(sheet) ? WorkbookOps.RowKeyToIndex(wbLocal.Worksheet(sheet), maxCol) : new(),
            RowBase = wbBase.Worksheets.Contains(sheet) ? WorkbookOps.RowKeyToIndex(wbBase.Worksheet(sheet), maxCol) : new(),
            RowRemote = wbRemote.Worksheets.Contains(sheet) ? WorkbookOps.RowKeyToIndex(wbRemote.Worksheet(sheet), maxCol) : new(),
        };
    }

    private static void EnsureHeaders(SheetChoiceCache cache)
    {
        var maxCol = cache.Out.LastColumnUsed()?.ColumnNumber() ?? 1;
        if (cache.HeaderMaxCol == maxCol)
        {
            return;
        }

        cache.HeaderOut = SheetSnapshot.HeaderIndex(SheetSnapshot.From(cache.Out, maxCol).Headers, true);
        cache.HeaderLocal = cache.Local is null
            ? new Dictionary<string, int>()
            : SheetSnapshot.HeaderIndex(SheetSnapshot.From(cache.Local, maxCol).Headers, true);
        cache.HeaderRemote = cache.Remote is null
            ? new Dictionary<string, int>()
            : SheetSnapshot.HeaderIndex(SheetSnapshot.From(cache.Remote, maxCol).Headers, true);
        cache.HeaderMaxCol = maxCol;
    }

    private static void MergeNewRowsClassic(string pathBase, string pathOther, string pathMerged)
    {
        using var wbBase = new XLWorkbook(pathBase);
        using var wbOther = new XLWorkbook(pathOther);
        using var wbOut = WorkbookOps.CreateEmpty();
        var names = UnionNames(wbBase, wbOther);
        foreach (var name in names)
        {
            var wsB = wbBase.Worksheets.Contains(name) ? wbBase.Worksheet(name) : null;
            var wsO = wbOther.Worksheets.Contains(name) ? wbOther.Worksheet(name) : null;
            if (wsB is null && wsO is null)
            {
                continue;
            }

            var maxCol = Math.Max(wsB?.LastColumnUsed()?.ColumnNumber() ?? 1, wsO?.LastColumnUsed()?.ColumnNumber() ?? 1);
            var rowsB = wsB is null ? new List<List<object?>>() : LoadNonEmptyKeyRows(wsB, maxCol);
            var rowsO = wsO is null ? new List<List<object?>>() : LoadNonEmptyKeyRows(wsO, maxCol);
            var baseRows = SheetSnapshot.RowsByRawKey(rowsB);
            var otherRows = SheetSnapshot.RowsByRawKey(rowsO);
            var baseOrdered = SheetSnapshot.OrderedRawKeys(rowsB);
            var otherOrdered = SheetSnapshot.OrderedRawKeys(rowsO);
            var newKeys = otherOrdered.Where(k => !baseRows.ContainsKey(k)).ToList();
            var merged = KeyNormalizer.MergeOrdered(baseOrdered, newKeys);
            var wsOut = wbOut.AddWorksheet(name);
            for (var r = 0; r < merged.Count; r++)
            {
                var key = merged[r];
                var row = baseRows.TryGetValue(key, out var fromBase) ? fromBase : otherRows[key];
                var isNew = newKeys.Contains(key);
                for (var c = 0; c < Math.Max(maxCol, row.Count); c++)
                {
                    var cell = wsOut.Cell(r + 1, c + 1);
                    var value = c < row.Count ? row[c] : null;
                    cell.Value = ToXl(value);
                    if (isNew)
                    {
                        cell.Style.Font.FontColor = NewFont;
                    }
                }
            }
        }

        WorkbookOps.EnsureSheet(wbOut);
        WorkbookOps.EnsureDirectory(pathMerged);
        wbOut.SaveAs(pathMerged);
    }

    private static void MergeNewColumnsClassic(string pathBase, string pathOther, string pathMerged)
    {
        using var wbBase = new XLWorkbook(pathBase);
        using var wbOther = new XLWorkbook(pathOther);
        using var wbOut = WorkbookOps.CreateEmpty();
        foreach (var name in UnionNames(wbBase, wbOther))
        {
            var wsB = wbBase.Worksheets.Contains(name) ? wbBase.Worksheet(name) : null;
            var wsO = wbOther.Worksheets.Contains(name) ? wbOther.Worksheet(name) : null;
            if (wsB is null && wsO is null)
            {
                continue;
            }

            var snapB = wsB is null ? SheetSnapshot.Empty(name) : SheetSnapshot.From(wsB);
            var snapO = wsO is null ? SheetSnapshot.Empty(name) : SheetSnapshot.From(wsO);
            var baseCols = snapB.Headers.Where(h => h.Length > 0).ToList();
            var otherCols = snapO.Headers.Where(h => h.Length > 0).ToList();
            var newCols = otherCols.Where(c => !baseCols.Contains(c)).ToList();
            var merged = KeyNormalizer.MergeOrdered(baseCols, newCols);
            var maxRow = Math.Max(snapB.MaxRow, snapO.MaxRow);
            var wsOut = wbOut.AddWorksheet(name);
            for (var c = 0; c < merged.Count; c++)
            {
                var colKey = merged[c];
                var isNew = newCols.Contains(colKey);
                var srcSnap = isNew ? snapO : snapB;
                var idx = srcSnap.Headers.FindIndex(h => h == colKey) + 1;
                for (var r = 1; r <= maxRow; r++)
                {
                    var cell = wsOut.Cell(r, c + 1);
                    var value = idx > 0 && r <= srcSnap.Rows.Count && idx <= srcSnap.Rows[r - 1].Count
                        ? srcSnap.Rows[r - 1][idx - 1]
                        : null;
                    cell.Value = ToXl(value);
                    if (isNew)
                    {
                        cell.Style.Font.FontColor = NewFont;
                    }
                }
            }
        }

        WorkbookOps.EnsureSheet(wbOut);
        WorkbookOps.EnsureDirectory(pathMerged);
        wbOut.SaveAs(pathMerged);
    }

    private static void MergeNewSheetsClassic(string pathBase, string pathOther, string pathMerged)
    {
        using var wbBase = new XLWorkbook(pathBase);
        using var wbOther = new XLWorkbook(pathOther);
        var baseSheets = wbBase.Worksheets.Select(ws => ws.Name).Where(n => !SheetFilter.ShouldSkip(n)).ToHashSet(StringComparer.Ordinal);
        var newSheets = wbOther.Worksheets.Select(ws => ws.Name).Where(n => !SheetFilter.ShouldSkip(n) && !baseSheets.Contains(n)).ToList();
        WorkbookOps.EnsureDirectory(pathMerged);
        if (newSheets.Count == 0)
        {
            File.Copy(pathBase, pathMerged, overwrite: true);
            return;
        }

        foreach (var name in newSheets)
        {
            WorkbookOps.CopyWorksheet(wbBase, wbOther.Worksheet(name), name);
        }

        WorkbookOps.EnsureSheet(wbBase);
        wbBase.SaveAs(pathMerged);
    }

    private static void ApplyChoicesClassic(
        string pathLocal,
        string pathRemote,
        string pathMerged,
        string pathBase,
        IReadOnlyList<MergeChoice> choices)
    {
        using var wbOut = WorkbookOps.CreateEmpty();
        using var wbLocal = new XLWorkbook(pathLocal);
        foreach (var ws in wbLocal.Worksheets.Where(ws => !SheetFilter.ShouldSkip(ws.Name)))
        {
            WorkbookOps.CopyWorksheet(wbOut, ws, ws.Name);
        }

        using var wbBase = new XLWorkbook(pathBase);
        using var wbRemote = new XLWorkbook(pathRemote);
        ApplyChoices(wbOut, wbLocal, wbBase, wbRemote, choices);
        WorkbookOps.EnsureSheet(wbOut);
        WorkbookOps.EnsureDirectory(pathMerged);
        wbOut.SaveAs(pathMerged);
    }

    private static List<string> UnionNames(XLWorkbook left, XLWorkbook right)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var names = new List<string>();
        foreach (var wb in new[] { left, right })
        {
            foreach (var ws in wb.Worksheets)
            {
                if (!SheetFilter.ShouldSkip(ws.Name) && seen.Add(ws.Name))
                {
                    names.Add(ws.Name);
                }
            }
        }

        return names;
    }

    private static List<List<object?>> LoadNonEmptyKeyRows(IXLWorksheet ws, int maxCol)
    {
        var snap = SheetSnapshot.From(ws, maxCol);
        return snap.Rows.Where(r => r.Count > 0 && CellText.From(r[0]).Length > 0).ToList();
    }

    private static XLCellValue ToXl(object? value) => value switch
    {
        null => Blank.Value,
        string s => s,
        bool b => b,
        DateTime dt => dt,
        sbyte or byte or short or ushort or int or uint or long or ulong or float or double or decimal =>
            Convert.ToDouble(value),
        _ => Convert.ToString(value) ?? "",
    };

    private sealed class SheetChoiceCache
    {
        public required IXLWorksheet Out { get; init; }
        public IXLWorksheet? Local { get; init; }
        public IXLWorksheet? Base { get; init; }
        public IXLWorksheet? Remote { get; init; }
        public required Dictionary<string, int> RowOut { get; init; }
        public required Dictionary<string, int> RowLocal { get; init; }
        public required Dictionary<string, int> RowBase { get; init; }
        public required Dictionary<string, int> RowRemote { get; init; }
        public Dictionary<string, int> HeaderOut { get; set; } = new();
        public Dictionary<string, int> HeaderLocal { get; set; } = new();
        public Dictionary<string, int> HeaderRemote { get; set; } = new();
        public int HeaderMaxCol { get; set; }
    }
}
