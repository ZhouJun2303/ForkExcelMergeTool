using System.Diagnostics;
using ExcelMergeFork.Core.Excel;

namespace ExcelMergeFork.Core.Merge;

public static class PreviewBuilder
{
    public static MergePreview Build(MergeSession session, MergeOptions options)
    {
        var started = Stopwatch.StartNew();
        var items = new List<PreviewItem>();
        var conflictEntries = new List<MergeChoice>();
        var summary = new PreviewSummary();
        var baseSession = options.BaseSide == "remote" ? session.Remote : session.Local;
        var otherSession = options.BaseSide == "remote" ? session.Local : session.Remote;
        var common = baseSession.SheetNames.Where(otherSession.HasSheet).ToList();

        if (!options.SkipNewRows)
        {
            foreach (var sheet in common)
            {
                var dataB = baseSession.Snapshot(sheet);
                var dataO = otherSession.Snapshot(sheet);
                var baseKeys = dataB.OrderedKeysNormalized.ToHashSet(StringComparer.Ordinal);
                foreach (var key in dataO.OrderedKeysNormalized)
                {
                    if (key.Length > 0 && !baseKeys.Contains(key))
                    {
                        items.Add(Make(session, sheet, key, "将新增行", PreviewTag.New));
                        summary.New++;
                    }
                }
            }
        }

        if (options.DeleteMissingRows)
        {
            foreach (var sheet in common)
            {
                var dataB = baseSession.Snapshot(sheet);
                var dataO = otherSession.Snapshot(sheet);
                var otherKeys = dataO.OrderedKeysNormalized.ToHashSet(StringComparer.Ordinal);
                foreach (var key in dataB.OrderedKeysNormalized)
                {
                    if (key.Length > 0 && !otherKeys.Contains(key))
                    {
                        items.Add(Make(session, sheet, key, "将删除行", PreviewTag.Delete));
                        summary.Delete++;
                    }
                }
            }
        }

        if (!options.SkipNewColumns)
        {
            foreach (var sheet in common)
            {
                var dataB = baseSession.Snapshot(sheet);
                var dataO = otherSession.Snapshot(sheet);
                var headerB = dataB.Headers.Where(h => h.Length > 0)
                    .Select(KeyNormalizer.HeaderForCompare)
                    .ToHashSet(StringComparer.Ordinal);
                foreach (var header in dataO.Headers.Where(h => h.Length > 0))
                {
                    if (!headerB.Contains(KeyNormalizer.HeaderForCompare(header)))
                    {
                        items.Add(Make(session, sheet, header, "将新增列", PreviewTag.New, column: true));
                        summary.New++;
                    }
                }
            }
        }

        if (options.DeleteMissingColumns)
        {
            foreach (var sheet in common)
            {
                var dataB = baseSession.Snapshot(sheet);
                var dataO = otherSession.Snapshot(sheet);
                var headerO = dataO.Headers.Where(h => h.Length > 0)
                    .Select(KeyNormalizer.HeaderForCompare)
                    .ToHashSet(StringComparer.Ordinal);
                var seen = new HashSet<string>(StringComparer.Ordinal);
                foreach (var header in dataB.Headers.Where(h => h.Length > 0))
                {
                    var norm = KeyNormalizer.HeaderForCompare(header);
                    if (seen.Add(norm) && !headerO.Contains(norm))
                    {
                        items.Add(Make(session, sheet, header, "将删除列", PreviewTag.Delete, column: true));
                        summary.Delete++;
                    }
                }
            }
        }

        if (options.AddNewSheets)
        {
            var baseSheets = baseSession.SheetNames.ToHashSet(StringComparer.Ordinal);
            foreach (var name in otherSession.SheetNames)
            {
                if (!baseSheets.Contains(name))
                {
                    items.Add(new PreviewItem { Sheet = name, Key = "新增 Sheet", Action = "将追加", Tag = PreviewTag.New });
                    summary.New++;
                }
            }
        }

        if (options.DeleteMissingSheets)
        {
            var otherSheets = otherSession.SheetNames.ToHashSet(StringComparer.Ordinal);
            foreach (var name in baseSession.AllSheetNames)
            {
                if (!otherSheets.Contains(name))
                {
                    items.Add(new PreviewItem { Sheet = name, Key = "删除 Sheet", Action = "将删除", Tag = PreviewTag.Delete });
                    summary.Delete++;
                }
            }
        }

        if (options.ResolveConflicts)
        {
            AppendConflicts(session, items, conflictEntries, summary);
            AppendAutoActions(session, items, summary);
        }

        return new MergePreview
        {
            Items = items,
            ConflictEntries = conflictEntries,
            Summary = summary,
            BaseSide = options.BaseSide,
            Options = options.ToLetterSet(),
            ElapsedMs = (int)started.ElapsedMilliseconds,
            SheetCount = session.UnionSheetNames().Count,
        };
    }

    private static void AppendConflicts(
        MergeSession session,
        List<PreviewItem> items,
        List<MergeChoice> conflictEntries,
        PreviewSummary summary)
    {
        foreach (var conflict in ConflictDetector.DetectRows(session))
        {
            switch (conflict.Type)
            {
                case ConflictType.AddLocal:
                    items.Add(Make(session, conflict.Sheet, conflict.Key, "信息：本地新增", PreviewTag.New, displayKey: $"{conflict.Key} (仅本地新增)"));
                    summary.Info++;
                    break;
                case ConflictType.AddRemote:
                    items.Add(Make(session, conflict.Sheet, conflict.Key, "信息：线上新增", PreviewTag.New, displayKey: $"{conflict.Key} (仅线上新增)"));
                    summary.Info++;
                    break;
                case ConflictType.AddConflict:
                    AddChoice(session, conflict, items, conflictEntries, summary, $"{conflict.Key} (新增冲突)", "将保留本地", PreviewTag.Conflict, "本地");
                    break;
                case ConflictType.DeleteConflictLocal:
                    AddChoice(session, conflict, items, conflictEntries, summary, $"{conflict.Key} (删除冲突：本地删)", "将保留线上（本地已删）", PreviewTag.DeleteConflict, "线上");
                    break;
                case ConflictType.DeleteConflictRemote:
                    AddChoice(session, conflict, items, conflictEntries, summary, $"{conflict.Key} (删除冲突：线上删)", "将保留本地（线上已删）", PreviewTag.DeleteConflict, "本地");
                    break;
                case ConflictType.ModifyConflict:
                    AddChoice(session, conflict, items, conflictEntries, summary, $"{conflict.Key} (修改冲突)", "将保留本地", PreviewTag.Conflict, "本地");
                    break;
            }
        }

        var existing = conflictEntries
            .Where(e => e.Kind == ConflictKind.Column)
            .Select(e => (e.Sheet, e.Key))
            .ToHashSet();
        foreach (var conflict in ConflictDetector.DetectColumns(session))
        {
            if (!existing.Add((conflict.Sheet, conflict.Key)))
            {
                continue;
            }

            AddChoice(
                session,
                conflict,
                items,
                conflictEntries,
                summary,
                $"{conflict.Key} (列冲突)",
                "将保留本地列",
                PreviewTag.Conflict,
                "本地");
        }
    }

    private static void AppendAutoActions(MergeSession session, List<PreviewItem> items, PreviewSummary summary)
    {
        var seen = items.Select(i => (i.Sheet, i.Key.Split(" (", 2)[0], i.Action)).ToHashSet();
        foreach (var action in ConflictDetector.DetectAutoActions(session))
        {
            var (label, tag) = action.Type switch
            {
                AutoActionType.TakeLocal => ("自动采用本地修改", PreviewTag.Modify),
                AutoActionType.TakeRemote => ("自动采用线上修改", PreviewTag.Modify),
                AutoActionType.DeleteLocal => ("自动删除（本地已删，线上未改）", PreviewTag.Delete),
                AutoActionType.DeleteRemote => ("自动删除（线上已删，本地未改）", PreviewTag.Delete),
                _ => ("自动合并", PreviewTag.Info),
            };
            if (!seen.Add((action.Sheet, action.Key, label)))
            {
                continue;
            }

            items.Add(Make(session, action.Sheet, action.Key, label, tag));
            if (tag == PreviewTag.Delete)
            {
                summary.Delete++;
            }
            else
            {
                summary.Info++;
            }
        }
    }

    private static void AddChoice(
        MergeSession session,
        ConflictItem conflict,
        List<PreviewItem> items,
        List<MergeChoice> conflictEntries,
        PreviewSummary summary,
        string key,
        string action,
        PreviewTag tag,
        string defaultChoice)
    {
        var idx = conflictEntries.Count;
        conflictEntries.Add(new MergeChoice
        {
            Sheet = conflict.Sheet,
            Key = conflict.Key,
            Choice = defaultChoice == "线上" ? "remote" : "local",
            Kind = conflict.Kind,
        });
        items.Add(Make(
            session,
            conflict.Sheet,
            conflict.Key,
            action,
            tag,
            column: conflict.Kind == ConflictKind.Column,
            displayKey: key,
            conflictIndex: idx));
        summary.Conflict++;
    }

    private static PreviewItem Make(
        MergeSession session,
        string sheet,
        string key,
        string action,
        PreviewTag tag,
        bool column = false,
        string? displayKey = null,
        int? conflictIndex = null)
    {
        return new PreviewItem
        {
            Sheet = sheet,
            Key = displayKey ?? key,
            Action = action,
            Tag = tag,
            ConflictIndex = conflictIndex,
            LocalValues = SideTexts(session.Local, sheet, key, column),
            RemoteValues = SideTexts(session.Remote, sheet, key, column),
            BaseValues = SideTexts(session.Base, sheet, key, column),
        };
    }

    public static IReadOnlyList<string> SideTexts(WorkbookSession workbook, string sheet, string key, bool column)
    {
        if (!workbook.Sheets.TryGetValue(sheet, out var snap))
        {
            return [];
        }

        if (column)
        {
            var idx = snap.Headers.FindIndex(h =>
                KeyNormalizer.HeaderForCompare(h) == KeyNormalizer.HeaderForCompare(key));
            return idx < 0 ? [] : FormatPairs(snap.Headers, snap.ColumnValues(idx + 1).Cast<object?>().ToList());
        }

        if (!snap.RowsByKey.TryGetValue(key, out var row) &&
            !snap.RowsByKey.TryGetValue(KeyNormalizer.Normalize(key), out row))
        {
            return [];
        }

        return FormatPairs(snap.Headers, row);
    }

    private static IReadOnlyList<string> FormatPairs(IReadOnlyList<string> headers, IReadOnlyList<object?> cells)
    {
        var lines = new List<string>(cells.Count);
        for (var i = 0; i < cells.Count; i++)
        {
            var header = i < headers.Count && headers[i].Length > 0 ? headers[i] : "列" + (i + 1);
            lines.Add(header + ": " + CellText.From(cells[i]));
        }

        return lines;
    }
}
