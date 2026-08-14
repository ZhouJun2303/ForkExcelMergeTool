namespace ExcelMergeFork.Core;

public static class AppConstants
{
    public const string LogFileName = "MergeExcelFork.log";
    public const string OptionsFileName = "merge_options.json";
    public const string BackupSubdir = "MergeExcelBackup";
    public const string CompareSuffix = "_compare";
    public const string SkipSheetPrefix = "#";
    public const string HomeEnvVar = "EXCEL_MERGE_FORK_HOME";
    public const string GitHubRepo = "ZhouJun2303/ForkExcelMergeTool";
    public const string UpdateAssetName = "ExcelMergeFork.exe";
    public const string UpdateSha256AssetName = "ExcelMergeFork.exe.sha256";
    public const int MergeOptionsSchema = 2;
    public const int FuzzyPrefixMinLen = 4;

    public const string ColorNew = "CCFFCC";
    public const string ColorChanged = "FFFF99";
    public const string ColorConflict = "FFCCCC";
    public const string FontNew = "008000";
    public const string FontModified = "CC6600";
}
