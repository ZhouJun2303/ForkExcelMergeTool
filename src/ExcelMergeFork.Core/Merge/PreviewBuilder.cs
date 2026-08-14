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
                        items.Add(Make(sheet, key, "将新增行", PreviewTag.New, dataB, dataO));
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
                        items.Add(Make(sheet, key, "将删除行", PreviewTag.Delete, dataB, dataO));
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
                        items.Add(new PreviewItem
                        {
                            Sheet = sheet,
                            Key = header,
                            Action = "将新增列",
                            Tag = PreviewTag.New,
                            LocalValues = [],
                            RemoteValues = [],
                            BaseValues = [],
                        });
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
                        items.Add(new PreviewItem
                        {
                            Sheet = sheet,
                            Key = header,
                            Action = "将删除列",
                            Tag = PreviewTag.Delete,
                        });
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
                    items.Add(Values(conflict, $"{conflict.Key} (仅本地新增)", "信息：本地新增", PreviewTag.New));
                    summary.Info++;
                    break;
                case ConflictType.AddRemote:
                    items.Add(Values(conflict, $"{conflict.Key} (仅线上新增)", "信息：线上新增", PreviewTag.New));
                    summary.Info++;
                    break;
                case ConflictType.AddConflict:
                    AddChoice(conflict, items, conflictEntries, summary, $"{conflict.Key} (新增冲突)", "将保留本地", PreviewTag.Conflict, "本地");
                    break;
                case ConflictType.DeleteConflictLocal:
                    AddChoice(conflict, items, conflictEntries, summary, $"{conflict.Key} (删除冲突：本地删)", "将保留线上（本地已删）", PreviewTag.DeleteConflict, "线上");
                    break;
                case ConflictType.DeleteConflictRemote:
                    AddChoice(conflict, items, conflictEntries, summary, $"{conflict.Key} (删除冲突：线上删)", "将保留本地（线上已删）", PreviewTag.DeleteConflict, "本地");
                    break;
                case ConflictType.ModifyConflict:
                    AddChoice(conflict, items, conflictEntries, summary, $"{conflict.Key} (修改冲突)", "将保留本地", PreviewTag.Conflict, "本地");
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

            items.Add(new PreviewItem
            {
                Sheet = action.Sheet,
                Key = action.Key,
                Action = label,
                Tag = tag,
            });
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
        items.Add(Values(conflict, key, action, tag, idx));
        summary.Conflict++;
    }

    private static PreviewItem Make(
        string sheet,
        string key,
        string action,
        PreviewTag tag,
        SheetSnapshot baseline,
        SheetSnapshot other)
    {
        baseline.RowsByKey.TryGetValue(key, out var left);
        other.RowsByKey.TryGetValue(key, out var right);
        return new PreviewItem
        {
            Sheet = sheet,
            Key = key,
            Action = action,
            Tag = tag,
            BaseValues = ToTexts(left),
            LocalValues = ToTexts(left),
            RemoteValues = ToTexts(right),
        };
    }

    private static PreviewItem Values(
        ConflictItem conflict,
        string key,
        string action,
        PreviewTag tag,
        int? index = null)
    {
        return new PreviewItem
        {
            Sheet = conflict.Sheet,
            Key = key,
            Action = action,
            Tag = tag,
            ConflictIndex = index,
            LocalValues = conflict.LocalCol ?? ToTexts(conflict.LocalRow),
            RemoteValues = conflict.RemoteCol ?? ToTexts(conflict.RemoteRow),
            BaseValues = conflict.BaseCol ?? ToTexts(conflict.BaseRow),
        };
    }

    private static IReadOnlyList<string> ToTexts(IReadOnlyList<object?>? row) =>
        row?.Select(CellText.From).ToList() ?? [];
}
