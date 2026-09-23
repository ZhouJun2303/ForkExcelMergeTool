using ExcelMergeFork.Core.Excel;

namespace ExcelMergeFork.Core.Git;

public sealed class RepoConflictExcel
{
    public required string RelativePath { get; init; }
    public required string WorktreePath { get; init; }
    public string? BaseBlob { get; init; }
    public string? LocalBlob { get; init; }
    public string? RemoteBlob { get; init; }
    public bool HasBothSides => !string.IsNullOrEmpty(LocalBlob) && !string.IsNullOrEmpty(RemoteBlob);
}

public sealed class RepoConflictScan
{
    public IReadOnlyList<RepoConflictExcel> Files { get; init; } = [];
    public string? Error { get; init; }
}

public static class RepoConflictExcelFinder
{
    public static RepoConflictScan Scan(string repoRoot)
    {
        if (string.IsNullOrWhiteSpace(repoRoot) || !Directory.Exists(repoRoot))
        {
            return new RepoConflictScan { Error = "仓库目录不存在" };
        }

        var result = GitRunner.Run(repoRoot, ["-c", "core.quotepath=false", "ls-files", "-u", "-z"], 30);
        if (result.ExitCode != 0)
        {
            var message = string.IsNullOrWhiteSpace(result.StdErr) ? "git ls-files 失败" : result.StdErr.Trim();
            return new RepoConflictScan { Error = message };
        }

        return new RepoConflictScan { Files = Parse(repoRoot, result.StdOut) };
    }

    public static IReadOnlyList<RepoConflictExcel> Parse(string repoRoot, string lsFilesOutput)
    {
        var repo = Path.GetFullPath(repoRoot);
        var grouped = new Dictionary<string, StageBlobs>(StringComparer.OrdinalIgnoreCase);
        foreach (var record in SplitRecords(lsFilesOutput))
        {
            if (!TryParseRecord(record, out var stage, out var blob, out var path))
            {
                continue;
            }

            if (!ExcelFormats.MergeDiffSupported(path) || !TryNormalizeRelative(path, out var relative))
            {
                continue;
            }

            if (!grouped.TryGetValue(relative, out var stages))
            {
                stages = new StageBlobs(relative);
                grouped[relative] = stages;
            }

            switch (stage)
            {
                case 1:
                    stages.BaseBlob = blob;
                    break;
                case 2:
                    stages.LocalBlob = blob;
                    break;
                case 3:
                    stages.RemoteBlob = blob;
                    break;
            }
        }

        var files = new List<RepoConflictExcel>();
        foreach (var stages in grouped.Values.OrderBy(item => item.RelativePath, StringComparer.OrdinalIgnoreCase))
        {
            if (!TryWorktreePath(repo, stages.RelativePath, out var worktree))
            {
                continue;
            }

            files.Add(new RepoConflictExcel
            {
                RelativePath = stages.RelativePath,
                WorktreePath = worktree,
                BaseBlob = stages.BaseBlob,
                LocalBlob = stages.LocalBlob,
                RemoteBlob = stages.RemoteBlob,
            });
        }

        return files;
    }

    private static IEnumerable<string> SplitRecords(string output)
    {
        if (string.IsNullOrEmpty(output))
        {
            yield break;
        }

        var parts = output.Contains('\0') ? output.Split('\0') : output.Split('\n');
        foreach (var part in parts)
        {
            var record = part.Trim('\r');
            if (record.Length > 0)
            {
                yield return record;
            }
        }
    }

    private static bool TryParseRecord(string record, out int stage, out string blob, out string path)
    {
        stage = 0;
        blob = "";
        path = "";
        var tab = record.IndexOf('\t');
        if (tab <= 0 || tab >= record.Length - 1)
        {
            return false;
        }

        var head = record[..tab].Split(' ');
        if (head.Length != 3 || !int.TryParse(head[2], out stage) || stage is < 1 or > 3)
        {
            return false;
        }

        blob = head[1];
        if (blob.Length is < 4 or > 64 || blob.Any(ch => !Uri.IsHexDigit(ch)))
        {
            return false;
        }

        path = record[(tab + 1)..];
        return path.Length > 0;
    }

    private static bool TryNormalizeRelative(string path, out string relative)
    {
        relative = path.Replace('\\', '/').Trim();
        if (relative.Length == 0 || Path.IsPathRooted(path) || relative.Contains(':', StringComparison.Ordinal))
        {
            return false;
        }

        var segments = relative.Split('/');
        if (segments.Any(segment => segment.Length == 0 || segment is "." or ".."))
        {
            return false;
        }

        relative = string.Join('/', segments);
        return true;
    }

    private static bool TryWorktreePath(string repoRoot, string relative, out string worktree)
    {
        var combined = Path.GetFullPath(Path.Combine(repoRoot, relative.Replace('/', Path.DirectorySeparatorChar)));
        var prefix = repoRoot.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
        if (!combined.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
        {
            worktree = "";
            return false;
        }

        worktree = combined;
        return true;
    }

    private sealed class StageBlobs(string relativePath)
    {
        public string RelativePath { get; } = relativePath;
        public string? BaseBlob { get; set; }
        public string? LocalBlob { get; set; }
        public string? RemoteBlob { get; set; }
    }
}
