namespace ExcelMergeFork.Core.Git;

public sealed class CompletionResult
{
    public bool Success { get; set; }
    public bool Staged { get; set; }
    public List<string> Cleaned { get; } = [];
    public List<string> Skipped { get; } = [];
    public List<string> Errors { get; } = [];
    public string Message { get; set; } = "";
}

public sealed class CleanupPolicy
{
    public List<string> AllowedRoots { get; }

    public CleanupPolicy(IEnumerable<string?> roots)
    {
        AllowedRoots = roots
            .Where(r => !string.IsNullOrWhiteSpace(r) && Directory.Exists(r))
            .Select(r => Path.GetFullPath(r!))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    public static CleanupPolicy Default() => new([
        Path.GetTempPath(),
        Environment.GetEnvironmentVariable("TEMP"),
        Environment.GetEnvironmentVariable("TMP"),
    ]);

    public bool Allows(string path)
    {
        if (!File.Exists(path))
        {
            return false;
        }

        var full = Path.GetFullPath(path);
        return AllowedRoots.Any(root => full.StartsWith(root.TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
                                        || string.Equals(Path.GetDirectoryName(full), root, StringComparison.OrdinalIgnoreCase));
    }
}

public static class GitCompletion
{
    public static CompletionResult StageAndCleanup(string pathMerged, string pathLocal, string pathBase, string pathRemote, CleanupPolicy? policy = null)
    {
        var result = new CompletionResult();
        var repo = GitRunner.FindWorktreeRoot(pathMerged);
        if (repo is null)
        {
            result.Errors.Add("无法确认 Git 仓库根目录");
            return result;
        }

        var absMerged = Path.GetFullPath(pathMerged);
        if (!absMerged.StartsWith(repo, StringComparison.OrdinalIgnoreCase))
        {
            result.Errors.Add("MERGED 不在当前 Git 仓库中，已停止确认流程: " + absMerged);
            return result;
        }

        var rel = Path.GetRelativePath(repo, absMerged).Replace('\\', '/');
        var add = GitRunner.Run(repo, ["add", "--", rel], 30);
        if (add.ExitCode != 0)
        {
            result.Errors.Add("git add 失败: " + FirstNonEmpty(add.StdErr, add.StdOut, "未知"));
            return result;
        }

        result.Staged = true;
        policy ??= CleanupPolicy.Default();
        foreach (var (label, path) in new[] { ("LOCAL", pathLocal), ("BASE", pathBase), ("REMOTE", pathRemote) })
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                continue;
            }

            if (policy.Allows(path))
            {
                try
                {
                    File.Delete(path);
                    result.Cleaned.Add(path);
                }
                catch (Exception ex)
                {
                    result.Errors.Add($"清理 {label} 失败: {ex.Message}");
                }
            }
            else
            {
                result.Skipped.Add(path);
            }
        }

        result.Success = result.Staged && result.Errors.Count == 0;
        result.Message = result.Success ? "已执行 git add，冲突已标记为已解决" : string.Join("; ", result.Errors);
        return result;
    }

    private static string FirstNonEmpty(params string[] values) =>
        values.Select(v => v.Trim()).FirstOrDefault(v => v.Length > 0) ?? "";
}
