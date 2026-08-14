using ClosedXML.Excel;

namespace ExcelMergeFork.Core.Excel;

public sealed class SheetSnapshot
{
    public required string Name { get; init; }
    public required int MaxColumn { get; init; }
    public required int MaxRow { get; init; }
    public required List<List<object?>> Rows { get; init; }
    public required List<int> RowIndices { get; init; }
    public required List<string> Headers { get; init; }
    public required Dictionary<string, List<object?>> RowsByKey { get; init; }
    public required Dictionary<string, int> KeyToRowIndex { get; init; }
    public required List<string> OrderedKeys { get; init; }
    public required List<string> OrderedKeysNormalized { get; init; }

    public HashSet<string> KeySet => RowsByKey.Keys.ToHashSet();

    public static SheetSnapshot From(IXLWorksheet worksheet, int? maxColumn = null)
    {
        var lastCol = Math.Max(1, maxColumn ?? worksheet.LastColumnUsed()?.ColumnNumber() ?? 1);
        var lastRow = Math.Max(1, worksheet.LastRowUsed()?.RowNumber() ?? 1);
        var rows = new List<List<object?>>(lastRow);
        var indices = new List<int>(lastRow);

        for (var r = 1; r <= lastRow; r++)
        {
            var row = new List<object?>(lastCol);
            for (var c = 1; c <= lastCol; c++)
            {
                row.Add(CellText.ToObject(worksheet.Cell(r, c).Value));
            }

            rows.Add(row);
            indices.Add(r);
        }

        ApplyMergedValues(worksheet, rows, lastCol);

        var headers = rows.Count > 0
            ? rows[0].Select(CellText.From).ToList()
            : [];

        var (byKey, keyToRow, ordered) = BuildKeyMaps(rows, indices);
        var orderedNorm = OrderedNormalized(rows);

        return new SheetSnapshot
        {
            Name = worksheet.Name,
            MaxColumn = lastCol,
            MaxRow = lastRow,
            Rows = rows,
            RowIndices = indices,
            Headers = headers,
            RowsByKey = byKey,
            KeyToRowIndex = keyToRow,
            OrderedKeys = ordered,
            OrderedKeysNormalized = orderedNorm,
        };
    }

    public static SheetSnapshot Empty(string name) => new()
    {
        Name = name,
        MaxColumn = 1,
        MaxRow = 1,
        Rows = [],
        RowIndices = [],
        Headers = [],
        RowsByKey = new Dictionary<string, List<object?>>(),
        KeyToRowIndex = new Dictionary<string, int>(),
        OrderedKeys = [],
        OrderedKeysNormalized = [],
    };

    public List<string> ColumnValues(int col1Based)
    {
        var values = new List<string>(Rows.Count);
        foreach (var row in Rows)
        {
            values.Add(col1Based <= row.Count ? CellText.From(row[col1Based - 1]) : "");
        }

        return values;
    }

    public static Dictionary<string, int> HeaderIndex(IReadOnlyList<string> headers, bool compareNormalize)
    {
        var map = new Dictionary<string, int>(StringComparer.Ordinal);
        for (var i = 0; i < headers.Count; i++)
        {
            var key = compareNormalize
                ? KeyNormalizer.HeaderForCompare(headers[i])
                : KeyNormalizer.Normalize(headers[i]);
            if (key.Length > 0 && !map.ContainsKey(key))
            {
                map[key] = i + 1;
            }
        }

        return map;
    }

    public static (Dictionary<string, List<object?>> rows, Dictionary<string, int> keyToRow, List<string> ordered)
        BuildKeyMaps(IReadOnlyList<List<object?>> rows, IReadOnlyList<int> indices)
    {
        var dict = new Dictionary<string, List<object?>>(StringComparer.Ordinal);
        var keyToRow = new Dictionary<string, int>(StringComparer.Ordinal);
        var ordered = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);

        for (var i = 0; i < rows.Count && i < indices.Count; i++)
        {
            var raw = rows[i].Count > 0 ? CellText.From(rows[i][0]) : "";
            var key = raw.Length == 0 ? $"__row_{indices[i]}" : KeyNormalizer.Normalize(raw);
            if (!seen.Add(key))
            {
                key = $"__row_{indices[i]}";
                seen.Add(key);
            }

            dict[key] = rows[i];
            keyToRow[key] = indices[i];
            ordered.Add(key);
        }

        return (dict, keyToRow, ordered);
    }

    public static List<string> OrderedNormalized(IReadOnlyList<List<object?>> rows)
    {
        var keys = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var row in rows)
        {
            var key = row.Count > 0 ? KeyNormalizer.Normalize(row[0]) : "";
            if (key.Length > 0 && seen.Add(key))
            {
                keys.Add(key);
            }
        }

        return keys;
    }

    public static List<string> OrderedRawKeys(IReadOnlyList<List<object?>> rows)
    {
        var keys = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var row in rows)
        {
            var key = row.Count > 0 ? CellText.From(row[0]) : "";
            if (key.Length > 0 && seen.Add(key))
            {
                keys.Add(key);
            }
        }

        return keys;
    }

    public static Dictionary<string, List<object?>> RowsByRawKey(IReadOnlyList<List<object?>> rows)
    {
        var dict = new Dictionary<string, List<object?>>(StringComparer.Ordinal);
        foreach (var row in rows)
        {
            var key = row.Count > 0 ? CellText.From(row[0]) : "";
            if (key.Length > 0)
            {
                dict[key] = row;
            }
        }

        return dict;
    }

    private static void ApplyMergedValues(IXLWorksheet worksheet, List<List<object?>> rows, int maxCol)
    {
        foreach (var range in worksheet.MergedRanges)
        {
            var value = CellText.ToObject(range.FirstCell().Value);
            if (value is null)
            {
                continue;
            }

            var minRow = range.RangeAddress.FirstAddress.RowNumber;
            var maxRow = range.RangeAddress.LastAddress.RowNumber;
            var minCol = range.RangeAddress.FirstAddress.ColumnNumber;
            var lastCol = range.RangeAddress.LastAddress.ColumnNumber;
            for (var r = minRow; r <= maxRow && r <= rows.Count; r++)
            {
                for (var c = minCol; c <= lastCol && c <= maxCol; c++)
                {
                    var current = rows[r - 1][c - 1];
                    if (current is null || (current is string s && string.IsNullOrWhiteSpace(s)))
                    {
                        rows[r - 1][c - 1] = value;
                    }
                }
            }
        }
    }
}
