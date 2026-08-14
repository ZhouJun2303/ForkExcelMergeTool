using ExcelMergeFork.Core.Excel;

namespace ExcelMergeFork.Core.Merge;

public static class ConflictDetector
{
    public static IReadOnlyList<ConflictItem> DetectRows(MergeSession session)
    {
        var conflicts = new List<ConflictItem>();
        foreach (var sheetName in session.UnionSheetNames())
        {
            var local = session.Local.Sheets.GetValueOrDefault(sheetName);
            var remote = session.Remote.Sheets.GetValueOrDefault(sheetName);
            var baseline = session.Base.Sheets.GetValueOrDefault(sheetName);
            var localRows = local?.RowsByKey ?? new Dictionary<string, List<object?>>();
            var remoteRows = remote?.RowsByKey ?? new Dictionary<string, List<object?>>();
            var baseRows = baseline?.RowsByKey ?? new Dictionary<string, List<object?>>();
            var keys = localRows.Keys.Concat(remoteRows.Keys).Concat(baseRows.Keys).ToHashSet(StringComparer.Ordinal);

            foreach (var key in keys)
            {
                localRows.TryGetValue(key, out var rowL);
                remoteRows.TryGetValue(key, out var rowR);
                baseRows.TryGetValue(key, out var rowB);

                if (rowB is null)
                {
                    if (rowL is not null && rowR is not null)
                    {
                        if (!CellText.RowsEqual(rowL, rowR))
                        {
                            conflicts.Add(new ConflictItem
                            {
                                Sheet = sheetName,
                                Key = key,
                                Type = ConflictType.AddConflict,
                                LocalRow = rowL,
                                RemoteRow = rowR,
                            });
                        }
                    }
                    else if (rowL is not null)
                    {
                        conflicts.Add(new ConflictItem
                        {
                            Sheet = sheetName,
                            Key = key,
                            Type = ConflictType.AddLocal,
                            LocalRow = rowL,
                            OnlyLocal = true,
                        });
                    }
                    else if (rowR is not null)
                    {
                        conflicts.Add(new ConflictItem
                        {
                            Sheet = sheetName,
                            Key = key,
                            Type = ConflictType.AddRemote,
                            RemoteRow = rowR,
                            OnlyRemote = true,
                        });
                    }
                }
                else if (rowL is null || rowR is null)
                {
                    if (rowL is null && rowR is null)
                    {
                        continue;
                    }

                    if (rowL is null)
                    {
                        conflicts.Add(new ConflictItem
                        {
                            Sheet = sheetName,
                            Key = key,
                            Type = ConflictType.DeleteConflictLocal,
                            RemoteRow = rowR,
                            BaseRow = rowB,
                        });
                    }
                    else
                    {
                        conflicts.Add(new ConflictItem
                        {
                            Sheet = sheetName,
                            Key = key,
                            Type = ConflictType.DeleteConflictRemote,
                            LocalRow = rowL,
                            BaseRow = rowB,
                        });
                    }
                }
                else if (!CellText.RowsEqual(rowL, rowR) &&
                         !CellText.RowsEqual(rowB, rowL) &&
                         !CellText.RowsEqual(rowB, rowR))
                {
                    conflicts.Add(new ConflictItem
                    {
                        Sheet = sheetName,
                        Key = key,
                        Type = ConflictType.ModifyConflict,
                        LocalRow = rowL,
                        RemoteRow = rowR,
                        BaseRow = rowB,
                    });
                }
            }
        }

        return conflicts;
    }

    public static IReadOnlyList<ConflictItem> DetectColumns(MergeSession session)
    {
        var conflicts = new List<ConflictItem>();
        foreach (var sheetName in session.UnionSheetNames())
        {
            if (!session.Local.HasSheet(sheetName) || !session.Remote.HasSheet(sheetName))
            {
                continue;
            }

            var local = session.Local.Snapshot(sheetName);
            var remote = session.Remote.Snapshot(sheetName);
            var baseline = session.Base.Sheets.GetValueOrDefault(sheetName);
            var maxCol = Math.Max(local.MaxColumn, remote.MaxColumn);
            if (baseline is not null)
            {
                maxCol = Math.Max(maxCol, baseline.MaxColumn);
            }

            var mapL = SheetSnapshot.HeaderIndex(local.Headers, compareNormalize: true);
            var mapR = SheetSnapshot.HeaderIndex(remote.Headers, compareNormalize: true);
            var mapB = baseline is null
                ? new Dictionary<string, int>()
                : SheetSnapshot.HeaderIndex(baseline.Headers, compareNormalize: true);
            var common = mapL.Keys.Intersect(mapR.Keys, StringComparer.Ordinal).ToList();
            var maxRow = Math.Max(local.MaxRow, remote.MaxRow);
            if (baseline is not null)
            {
                maxRow = Math.Max(maxRow, baseline.MaxRow);
            }

            foreach (var headerNorm in common)
            {
                var colL = Pad(local.ColumnValues(mapL[headerNorm]), maxRow);
                var colR = Pad(remote.ColumnValues(mapR[headerNorm]), maxRow);
                var colB = mapB.TryGetValue(headerNorm, out var bIdx) && baseline is not null
                    ? Pad(baseline.ColumnValues(bIdx), maxRow)
                    : [];

                if (colB.Count > 0 && (CellText.ColumnsEqual(colL, colB) || CellText.ColumnsEqual(colR, colB)))
                {
                    continue;
                }

                if (!CellText.ColumnsEqual(colL, colR))
                {
                    var display = local.Headers.ElementAtOrDefault(mapL[headerNorm] - 1);
                    if (string.IsNullOrEmpty(display))
                    {
                        display = remote.Headers.ElementAtOrDefault(mapR[headerNorm] - 1);
                    }

                    conflicts.Add(new ConflictItem
                    {
                        Sheet = sheetName,
                        Key = display ?? headerNorm,
                        Type = ConflictType.ColumnConflict,
                        Kind = ConflictKind.Column,
                        LocalCol = colL,
                        RemoteCol = colR,
                        BaseCol = colB,
                    });
                }
            }
        }

        return conflicts;
    }

    public static IReadOnlyList<AutoRowAction> DetectAutoActions(MergeSession session)
    {
        var actions = new List<AutoRowAction>();
        foreach (var sheetName in session.UnionSheetNames())
        {
            var local = session.Local.Sheets.GetValueOrDefault(sheetName);
            var remote = session.Remote.Sheets.GetValueOrDefault(sheetName);
            var baseline = session.Base.Sheets.GetValueOrDefault(sheetName);
            var dictL = local?.RowsByKey ?? new Dictionary<string, List<object?>>();
            var dictR = remote?.RowsByKey ?? new Dictionary<string, List<object?>>();
            var dictB = baseline?.RowsByKey ?? new Dictionary<string, List<object?>>();
            var mapL = local?.KeyToRowIndex ?? new Dictionary<string, int>();
            var mapR = remote?.KeyToRowIndex ?? new Dictionary<string, int>();
            var mapB = baseline?.KeyToRowIndex ?? new Dictionary<string, int>();

            foreach (var key in dictL.Keys.Concat(dictR.Keys).Concat(dictB.Keys).ToHashSet(StringComparer.Ordinal))
            {
                if (mapL.GetValueOrDefault(key) == 1 ||
                    mapR.GetValueOrDefault(key) == 1 ||
                    mapB.GetValueOrDefault(key) == 1)
                {
                    continue;
                }

                dictB.TryGetValue(key, out var rowB);
                if (rowB is null)
                {
                    continue;
                }

                dictL.TryGetValue(key, out var rowL);
                dictR.TryGetValue(key, out var rowR);
                if (rowL is null && rowR is null)
                {
                    continue;
                }

                if (rowL is null)
                {
                    if (CellText.RowsEqual(rowR, rowB))
                    {
                        actions.Add(new AutoRowAction
                        {
                            Sheet = sheetName,
                            Key = key,
                            Choice = "local",
                            Type = AutoActionType.DeleteLocal,
                        });
                    }

                    continue;
                }

                if (rowR is null)
                {
                    if (CellText.RowsEqual(rowL, rowB))
                    {
                        actions.Add(new AutoRowAction
                        {
                            Sheet = sheetName,
                            Key = key,
                            Choice = "remote",
                            Type = AutoActionType.DeleteRemote,
                        });
                    }

                    continue;
                }

                if (CellText.RowsEqual(rowL, rowR))
                {
                    continue;
                }

                var localChanged = !CellText.RowsEqual(rowL, rowB);
                var remoteChanged = !CellText.RowsEqual(rowR, rowB);
                if (localChanged && !remoteChanged)
                {
                    actions.Add(new AutoRowAction
                    {
                        Sheet = sheetName,
                        Key = key,
                        Choice = "local",
                        Type = AutoActionType.TakeLocal,
                    });
                }
                else if (remoteChanged && !localChanged)
                {
                    actions.Add(new AutoRowAction
                    {
                        Sheet = sheetName,
                        Key = key,
                        Choice = "remote",
                        Type = AutoActionType.TakeRemote,
                    });
                }
            }
        }

        return actions;
    }

    public static List<MergeChoice> MergeChoices(IEnumerable<MergeChoice>? userChoices, IEnumerable<AutoRowAction>? autoActions)
    {
        var merged = new List<MergeChoice>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var item in (userChoices ?? []).Concat((autoActions ?? []).Select(a => a.ToChoice())))
        {
            if (seen.Add(item.ChoiceKey))
            {
                merged.Add(item);
            }
        }

        return merged;
    }

    private static List<string> Pad(List<string> values, int maxRow)
    {
        while (values.Count < maxRow)
        {
            values.Add("");
        }

        return values;
    }
}
