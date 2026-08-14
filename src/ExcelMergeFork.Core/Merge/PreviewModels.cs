namespace ExcelMergeFork.Core.Merge;

public enum PreviewTag
{
    Info,
    New,
    Delete,
    Modify,
    DeleteConflict,
    Conflict,
}

public sealed class PreviewItem
{
    public required string Sheet { get; init; }
    public required string Key { get; init; }
    public required string Action { get; set; }
    public required PreviewTag Tag { get; init; }
    public int? ConflictIndex { get; init; }
    public IReadOnlyList<string> LocalValues { get; init; } = [];
    public IReadOnlyList<string> RemoteValues { get; init; } = [];
    public IReadOnlyList<string> BaseValues { get; init; } = [];
}

public sealed class PreviewSummary
{
    public int New { get; set; }
    public int Delete { get; set; }
    public int Conflict { get; set; }
    public int Info { get; set; }
}

public sealed class MergePreview
{
    public required IReadOnlyList<PreviewItem> Items { get; init; }
    public required IReadOnlyList<MergeChoice> ConflictEntries { get; init; }
    public required PreviewSummary Summary { get; init; }
    public required string BaseSide { get; init; }
    public required IReadOnlyCollection<string> Options { get; init; }
    public int ElapsedMs { get; init; }
    public int SheetCount { get; init; }
}
