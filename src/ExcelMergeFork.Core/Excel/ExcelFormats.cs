namespace ExcelMergeFork.Core.Excel;

public static class ExcelFormats
{
    public static readonly string[] CommonExtensions =
    [
        "xls", "xlsx", "xlsm", "xlsb", "xlt", "xltx", "xltm", "xla", "xlam", "xlw",
    ];

    public static readonly string[] MergeDiffExtensions = ["xlsx", "xltx"];

    public static string CommonExtensionText => string.Join(", ", CommonExtensions.Select(e => "." + e));

    public static string MergeDiffExtensionText => string.Join(", ", MergeDiffExtensions.Select(e => "." + e));

    public static string NormalizedExt(string? path)
    {
        var ext = Path.GetExtension(path ?? "");
        return ext.StartsWith('.') ? ext[1..].ToLowerInvariant() : ext.ToLowerInvariant();
    }

    public static bool MergeDiffSupported(string? path) =>
        MergeDiffExtensions.Contains(NormalizedExt(path));

    public static IReadOnlyList<string> GitAttrLines()
    {
        var lines = new List<string> { "# ExcelMergeFork managed entry" };
        foreach (var ext in CommonExtensions)
        {
            var pattern = string.Concat(ext.Select(c => $"[{char.ToLowerInvariant(c)}{char.ToUpperInvariant(c)}]"));
            lines.Add($"*.{pattern} merge=excelmergefork");
        }
        return lines;
    }
}
