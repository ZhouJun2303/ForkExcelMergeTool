using ExcelMergeFork.Core.Git;
using ExcelMergeFork.Core.Routing;

namespace ExcelMergeFork.Api;

public sealed class BootFile
{
    public required string Role { get; init; }
    public required string Name { get; init; }
    public required string Size { get; init; }
    public required string Path { get; init; }
}

public sealed class DesktopContext
{
    public LaunchRequest Request { get; init; } = new() { Mode = LaunchMode.Settings };
    public string Screen { get; set; } = "settings";
    public string? MergeId { get; set; }
    public string? CompareId { get; set; }
    public GitDriverRequest? Driver { get; set; }
    public bool WriteBackSucceeded { get; set; }
    public int ExitCode { get; set; }
    public string? Error { get; set; }
    public string ChoiceTitle { get; set; } = "";
    public string ChoiceNotice { get; set; } = "";
    public string ChoiceWarn { get; set; } = "";
    public bool CanMerge { get; set; } = true;
    public List<BootFile> ChoiceFiles { get; set; } = [];
    public bool BackupOk { get; set; }
    public string BackupHeadline { get; set; } = "";
    public string BackupDetail { get; set; } = "";
    public string? BackupDir { get; set; }
    public List<BootFile> BackupFiles { get; set; } = [];
}
