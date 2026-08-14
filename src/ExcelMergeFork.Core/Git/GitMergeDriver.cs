using ExcelMergeFork.Core.Merge;

namespace ExcelMergeFork.Core.Git;

public sealed class GitDriverRequest
{
    public required string BasePath { get; init; }
    public required string CurrentPath { get; init; }
    public required string OtherPath { get; init; }
    public required string RepoPath { get; init; }
    public required string WorkDir { get; init; }
    public required string LocalCopy { get; init; }
    public required string BaseCopy { get; init; }
    public required string RemoteCopy { get; init; }
    public required string MergedCopy { get; init; }
}

public static class GitMergeDriver
{
    public static GitDriverRequest Prepare(string basePath, string currentPath, string otherPath, string repoPath)
    {
        var work = Directory.CreateTempSubdirectory("excelmergefork_driver_");
        var localCopy = Path.Combine(work.FullName, "local" + Path.GetExtension(currentPath));
        var baseCopy = Path.Combine(work.FullName, "base" + Path.GetExtension(basePath));
        var remoteCopy = Path.Combine(work.FullName, "remote" + Path.GetExtension(otherPath));
        var mergedCopy = Path.Combine(work.FullName, "merged" + Path.GetExtension(currentPath));
        File.Copy(currentPath, localCopy, true);
        File.Copy(basePath, baseCopy, true);
        File.Copy(otherPath, remoteCopy, true);
        return new GitDriverRequest
        {
            BasePath = Path.GetFullPath(basePath),
            CurrentPath = Path.GetFullPath(currentPath),
            OtherPath = Path.GetFullPath(otherPath),
            RepoPath = repoPath,
            WorkDir = work.FullName,
            LocalCopy = localCopy,
            BaseCopy = baseCopy,
            RemoteCopy = remoteCopy,
            MergedCopy = mergedCopy,
        };
    }

    public static CompletionResult WriteBack(GitDriverRequest request)
    {
        var result = new CompletionResult();
        if (!File.Exists(request.MergedCopy))
        {
            result.Errors.Add("合并结果不存在: " + request.MergedCopy);
            return result;
        }

        var targetDir = Path.GetDirectoryName(request.CurrentPath) ?? ".";
        Directory.CreateDirectory(targetDir);
        var tmp = Path.Combine(targetDir, ".excelmergefork_driver_" + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            File.Copy(request.MergedCopy, tmp, true);
            File.Copy(tmp, request.CurrentPath, true);
            result.Success = true;
            result.Message = "Git merge driver 已写回当前文件，Git 将继续合并。";
        }
        catch (Exception ex)
        {
            result.Errors.Add("写回 Git 当前文件失败: " + ex.Message);
        }
        finally
        {
            try { File.Delete(tmp); } catch { }
        }

        return result;
    }

    public static int WindowCloseExitCode(bool writeBackSucceeded) => writeBackSucceeded ? 0 : 1;

    public static string ContextPath(string currentPath, string repoPath)
    {
        if (string.IsNullOrWhiteSpace(repoPath))
        {
            return currentPath;
        }

        try
        {
            var repo = Path.GetFullPath(repoPath);
            if (Directory.Exists(repo))
            {
                return Path.Combine(repo, Path.GetFileName(currentPath));
            }

            return repo;
        }
        catch
        {
            return currentPath;
        }
    }
}
