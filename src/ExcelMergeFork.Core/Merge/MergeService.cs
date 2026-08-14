using ExcelMergeFork.Core.Backup;

namespace ExcelMergeFork.Core.Merge;

public sealed class MergeResult
{
    public int ExitCode { get; init; }
    public BackupInfo? Backup { get; init; }
    public string? Error { get; init; }
}

public static class MergeService
{
    public static MergeResult Run(
        string pathLocal,
        string pathBase,
        string pathRemote,
        string pathMerged,
        string? mode = null,
        MergeOptions? options = null,
        IReadOnlyList<MergeChoice>? choices = null,
        string baseSide = "local",
        string? backupRoot = null,
        string? backupContextPath = null,
        bool createBackup = true)
    {
        if (!File.Exists(pathLocal) || !File.Exists(pathBase) || !File.Exists(pathRemote))
        {
            return new MergeResult { ExitCode = 1, Error = "参数或文件错误：LOCAL/BASE/REMOTE 不存在" };
        }

        try
        {
            if (options is not null)
            {
                MergePipeline.RunByOptions(pathLocal, pathBase, pathRemote, pathMerged, options, choices);
            }
            else
            {
                MergePipeline.RunByMode(pathLocal, pathBase, pathRemote, pathMerged, mode ?? "E", baseSide, choices);
            }
        }
        catch (Exception ex)
        {
            AppLog.Exception("合并异常", ex);
            return new MergeResult { ExitCode = 2, Error = ex.Message };
        }

        if (!createBackup)
        {
            AppLog.Info("合并完成 MERGED=" + pathMerged);
            return new MergeResult { ExitCode = 0 };
        }

        try
        {
            var backup = BackupService.CreateMergeBackup(pathLocal, pathRemote, pathMerged, backupRoot, backupContextPath);
            AppLog.Info($"合并完成 MERGED={pathMerged} 备份={backup.Dir}");
            return new MergeResult { ExitCode = 0, Backup = backup };
        }
        catch (Exception ex)
        {
            AppLog.Exception("备份异常", ex);
            return new MergeResult { ExitCode = 2, Error = "备份失败。" + ex.Message };
        }
    }
}
