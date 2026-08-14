namespace ExcelMergeFork.Core.Compare;

public sealed class DiffRow
{
    public required string Sheet { get; init; }
    public required string Key { get; init; }
    public required string Status { get; init; }
    public required string Left { get; init; }
    public required string Right { get; init; }
}

public sealed class CompareResult
{
    public required string OutputPath { get; init; }
    public required IReadOnlyList<string> SheetNames { get; init; }
    public required IReadOnlyList<DiffRow> Rows { get; init; }
    public int ElapsedMs { get; init; }
}
