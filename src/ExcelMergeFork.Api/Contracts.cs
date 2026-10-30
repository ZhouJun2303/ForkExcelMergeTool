namespace ExcelMergeFork.Api;

public sealed class OpenMergeRequest
{
    public string? Local { get; set; }
    public string? Base { get; set; }
    public string? Remote { get; set; }
    public string? Merged { get; set; }
    public string? Mode { get; set; }
    public MergeOptionsBody? Options { get; set; }
}

public sealed class MergeOptionsBody
{
    public bool SkipNewRows { get; set; }
    public bool SkipNewColumns { get; set; }
    public bool DeleteMissingRows { get; set; }
    public bool DeleteMissingColumns { get; set; }
    public bool AddNewSheets { get; set; } = true;
    public bool DeleteMissingSheets { get; set; }
    public bool ResolveConflicts { get; set; } = true;
    public string? BaseSide { get; set; }
    public string? Mode { get; set; }
}

public sealed class ChoiceRequest
{
    public string? Sheet { get; set; }
    public string? Key { get; set; }
    public string? Kind { get; set; }
    public string? Choice { get; set; }
}

public sealed class SelectSheetRequest
{
    public int Index { get; set; }
}

public sealed class GenerateRequest
{
    public string? Merged { get; set; }
    public bool CreateBackup { get; set; }
}

public sealed class OpenCompareRequest
{
    public string? Left { get; set; }
    public string? Right { get; set; }
    public bool IncludeSame { get; set; }
}

public sealed class CompareOptionsBody
{
    public bool IncludeSame { get; set; }
}

public sealed class SettingsBody
{
    public bool? SkipNewRows { get; set; }
    public bool? SkipNewColumns { get; set; }
    public bool? DeleteMissingRows { get; set; }
    public bool? DeleteMissingColumns { get; set; }
    public bool? AddNewSheets { get; set; }
    public bool? DeleteMissingSheets { get; set; }
    public bool? ResolveConflicts { get; set; }
    public string? StartupFeature { get; set; }
    public string? BackupRootDir { get; set; }
    public bool? AutoOpenMerged { get; set; }
    public bool? AutoOpenCompare { get; set; }
    public bool? DarkTheme { get; set; }
    public Dictionary<string, bool>? DiffFilter { get; set; }
}

public sealed class ChooseBody
{
    public string? Choice { get; set; }
}

public sealed class OpenPathRequest
{
    public string? Path { get; set; }
}

public sealed class AiCheckRequest
{
    public bool? All { get; set; }
    public string? Path { get; set; }
    public bool Checked { get; set; }
}

public sealed class AiPromptRequest
{
    public bool Detail { get; set; }
}

public readonly record struct PageQuery(string? Q, int Offset, int Limit)
{
    public static PageQuery From(HttpRequest request)
    {
        var q = request.Query["q"].ToString();
        var offset = int.TryParse(request.Query["offset"], out var parsedOffset) ? Math.Max(0, parsedOffset) : 0;
        var limit = int.TryParse(request.Query["limit"], out var parsedLimit) ? Math.Clamp(parsedLimit, 1, 500) : 150;
        return new PageQuery(string.IsNullOrWhiteSpace(q) ? null : q, offset, limit);
    }
}
