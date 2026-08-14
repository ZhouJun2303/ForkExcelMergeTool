using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using ExcelMergeFork.Core.Git;

namespace ExcelMergeFork.Core.Backup;

public sealed class BackupInfo
{
    public required string Dir { get; init; }
    public required string Time { get; init; }
    public required Dictionary<string, string> Files { get; init; }
    public string Label { get; init; } = "";
    public string Root { get; init; } = "";
    public string Project { get; init; } = "";
    public string? Local { get; init; }
    public string? Remote { get; init; }
    public string? Merged { get; init; }
}

public static class BackupService
{
    private static readonly Regex InvalidDirChars = new(@"[<>:""/\\|?*\x00-\x1f]", RegexOptions.Compiled);

    public static string ResolveRoot(string contextPath, string? backupRoot = null)
    {
        var root = (backupRoot ?? "").Trim();
        if (root.Length == 0)
        {
            root = Settings.AppSettingsStore.Load().BackupRootDir;
        }

        if (root.Length == 0)
        {
            var mergedDir = Path.GetDirectoryName(Path.GetFullPath(contextPath)) ?? ".";
            root = Path.Combine(mergedDir, AppConstants.BackupSubdir);
        }

        return Path.GetFullPath(root);
    }

    public static string ProjectName(string contextPath)
    {
        var mergedDir = Path.GetDirectoryName(Path.GetFullPath(contextPath)) ?? Directory.GetCurrentDirectory();
        var projectDir = GitRunner.FindWorktreeRoot(mergedDir) ?? mergedDir;
        return Sanitize(Path.GetFileName(projectDir.TrimEnd(Path.DirectorySeparatorChar)), "Project");
    }

    public static BackupInfo CreateMergeBackup(string pathLocal, string pathRemote, string pathMerged, string? backupRoot = null, string? contextPath = null)
    {
        var context = contextPath ?? pathMerged;
        var dir = CreateNamedDir(context, backupRoot, out var stamp, out var label, out var localInfo, out var remoteInfo);
        var local = BackupPath(dir, context, "local", localInfo, remoteInfo);
        var remote = BackupPath(dir, context, "remote", localInfo, remoteInfo);
        var merged = BackupPath(dir, context, "merged", localInfo, remoteInfo);
        File.Copy(pathLocal, local, overwrite: true);
        File.Copy(pathRemote, remote, overwrite: true);
        File.Copy(pathMerged, merged, overwrite: true);
        return new BackupInfo
        {
            Dir = dir,
            Time = stamp,
            Label = label,
            Root = ResolveRoot(context, backupRoot),
            Project = ProjectName(context),
            Local = local,
            Remote = remote,
            Merged = merged,
            Files = new Dictionary<string, string>
            {
                ["local"] = local,
                ["remote"] = remote,
                ["merged"] = merged,
            },
        };
    }

    public static BackupInfo CreateQuickBackup(IEnumerable<(string Label, string Path)> files, string contextPath, string? backupRoot = null)
    {
        var dir = CreateNamedDir(contextPath, backupRoot, out var stamp, out var label, out var localInfo, out var remoteInfo);
        var copied = new Dictionary<string, string>(StringComparer.Ordinal);
        var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var (rawLabel, src) in files)
        {
            if (string.IsNullOrWhiteSpace(src) || !File.Exists(src))
            {
                continue;
            }

            var commitLabel = rawLabel switch
            {
                "current" => "local",
                "other" => "remote",
                _ => rawLabel,
            };
            string dest;
            if (commitLabel is "local" or "remote" or "merged")
            {
                dest = BackupPath(dir, contextPath, commitLabel, localInfo, remoteInfo);
            }
            else
            {
                dest = Path.Combine(dir, $"{Sanitize(Path.GetFileNameWithoutExtension(src), commitLabel, 72)}_{Sanitize(commitLabel, "backup", 24)}{Path.GetExtension(src)}");
            }

            var candidate = Path.GetFileName(dest);
            var i = 2;
            while (!used.Add(candidate))
            {
                candidate = $"{Path.GetFileNameWithoutExtension(dest)}_{i:00}{Path.GetExtension(dest)}";
                i++;
            }

            dest = Path.Combine(dir, candidate);
            File.Copy(src, dest, overwrite: true);
            copied[rawLabel] = dest;
        }

        if (copied.Count == 0)
        {
            throw new FileNotFoundException("没有可备份的输入文件");
        }

        return new BackupInfo
        {
            Dir = dir,
            Time = stamp,
            Label = label,
            Files = copied,
        };
    }

    private static string CreateNamedDir(
        string contextPath,
        string? backupRoot,
        out string stamp,
        out string label,
        out GitCommitInfo? localInfo,
        out GitCommitInfo? remoteInfo)
    {
        var parent = Path.Combine(ResolveRoot(contextPath, backupRoot), ProjectName(contextPath));
        Directory.CreateDirectory(parent);
        (localInfo, remoteInfo) = GitRunner.MergeInfo(contextPath);
        label = BackupLabel(contextPath, localInfo, remoteInfo);
        stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        var baseName = string.IsNullOrEmpty(label) ? stamp : $"{stamp}__{label}";
        foreach (var name in new[] { baseName }.Concat(Enumerable.Range(2, 20).Select(i => $"{baseName}_{i:00}")))
        {
            var dir = Path.Combine(parent, name);
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
                stamp = name;
                return dir;
            }
        }

        var fallback = DateTime.Now.ToString("yyyyMMdd_HHmmss_ffffff");
        var path = Path.Combine(parent, fallback);
        Directory.CreateDirectory(path);
        stamp = fallback;
        return path;
    }

    private static string BackupLabel(string contextPath, GitCommitInfo? local, GitCommitInfo? remote)
    {
        var excel = Sanitize(Path.GetFileNameWithoutExtension(contextPath), "Excel", 42);
        return Sanitize($"{excel}__{CommitPart(local, "L", true)}__{CommitPart(remote, "R", true)}", "ExcelBackup", 96);
    }

    private static string BackupPath(string backupDir, string contextPath, string label, GitCommitInfo? local, GitCommitInfo? remote)
    {
        var excel = Sanitize(Path.GetFileNameWithoutExtension(contextPath), "Excel", 54);
        var ext = Path.GetExtension(contextPath);
        if (string.IsNullOrEmpty(ext))
        {
            ext = ".xlsx";
        }

        var commit = label switch
        {
            "local" => CommitPart(local, "L", true),
            "remote" => CommitPart(remote, "R", true),
            "merged" => $"{CommitPart(local, "L", false)}__{CommitPart(remote, "R", false)}",
            _ => Sanitize(label, "backup", 24),
        };
        var stem = Sanitize($"{excel}__{label}__{commit}", $"{excel}_{label}", 96);
        var full = Path.Combine(backupDir, stem + ext);
        if (Path.GetFullPath(full).Length <= 240)
        {
            return full;
        }

        var digest = Convert.ToHexString(SHA1.HashData(Encoding.UTF8.GetBytes(Path.GetFullPath(full))))[..10].ToLowerInvariant();
        return Path.Combine(backupDir, $"{Sanitize(label, "backup", 18)}__{digest}{ext}");
    }

    private static string CommitPart(GitCommitInfo? info, string prefix, bool includeMessage)
    {
        if (info is null || string.IsNullOrEmpty(info.Hash))
        {
            return $"{prefix}-unknown-nohash-nomsg";
        }

        var author = Sanitize(info.Author, "unknown", 28);
        var hash = Sanitize(info.ShortHash, "nohash", 16);
        if (!includeMessage)
        {
            return $"{prefix}-{author}-{hash}";
        }

        return $"{prefix}-{author}-{hash}-{Sanitize(info.Message, "nomsg", 54)}";
    }

    private static string Sanitize(string? name, string fallback, int? maxLen = null)
    {
        var cleaned = InvalidDirChars.Replace(name ?? "", "_");
        cleaned = Regex.Replace(cleaned, @"\s+", " ").Trim(' ', '.', '_');
        if (maxLen is int n && cleaned.Length > n)
        {
            cleaned = cleaned[..n].TrimEnd(' ', '.', '_');
        }

        return string.IsNullOrWhiteSpace(cleaned) ? fallback : cleaned;
    }
}
