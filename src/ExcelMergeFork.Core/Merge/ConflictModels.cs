namespace ExcelMergeFork.Core.Merge;

public enum ConflictKind
{
    Row,
    Column,
}

public enum ConflictType
{
    AddLocal,
    AddRemote,
    AddConflict,
    DeleteConflictLocal,
    DeleteConflictRemote,
    ModifyConflict,
    ColumnConflict,
}

public enum AutoActionType
{
    TakeLocal,
    TakeRemote,
    DeleteLocal,
    DeleteRemote,
}

public sealed class ConflictItem
{
    public required string Sheet { get; init; }
    public required string Key { get; init; }
    public required ConflictType Type { get; init; }
    public ConflictKind Kind { get; init; } = ConflictKind.Row;
    public IReadOnlyList<object?>? LocalRow { get; init; }
    public IReadOnlyList<object?>? RemoteRow { get; init; }
    public IReadOnlyList<object?>? BaseRow { get; init; }
    public IReadOnlyList<string>? LocalCol { get; init; }
    public IReadOnlyList<string>? RemoteCol { get; init; }
    public IReadOnlyList<string>? BaseCol { get; init; }
    public bool OnlyLocal { get; init; }
    public bool OnlyRemote { get; init; }
}

public sealed class MergeChoice
{
    public required string Sheet { get; set; }
    public required string Key { get; set; }
    public string Choice { get; set; } = "local";
    public ConflictKind Kind { get; set; } = ConflictKind.Row;
    public string? AutoType { get; set; }

    public string ChoiceKey => $"{Sheet}\u0001{Key}\u0001{Kind}";
}

public sealed class AutoRowAction
{
    public required string Sheet { get; init; }
    public required string Key { get; init; }
    public required string Choice { get; init; }
    public required AutoActionType Type { get; init; }

    public MergeChoice ToChoice() => new()
    {
        Sheet = Sheet,
        Key = Key,
        Choice = Choice,
        Kind = ConflictKind.Row,
        AutoType = Type switch
        {
            AutoActionType.TakeLocal => "take_local",
            AutoActionType.TakeRemote => "take_remote",
            AutoActionType.DeleteLocal => "delete_local",
            AutoActionType.DeleteRemote => "delete_remote",
            _ => null,
        },
    };
}
