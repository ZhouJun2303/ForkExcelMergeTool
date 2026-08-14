namespace ExcelMergeFork.Core.Excel;

public static class SheetFilter
{
    public static bool ShouldSkip(string? name) =>
        (name ?? "").TrimStart().StartsWith(AppConstants.SkipSheetPrefix, StringComparison.Ordinal);
}
