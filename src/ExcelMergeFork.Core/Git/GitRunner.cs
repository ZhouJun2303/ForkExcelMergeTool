using System.Diagnostics;
using System.Text;

namespace ExcelMergeFork.Core.Git;

public sealed class GitCommitInfo
{
    public string Hash { get; init; } = "";
    public string ShortHash { get; init; } = "";
    public string Author { get; init; } = "";
    public string Email { get; init; } = "";
    public string Date { get; init; } = "";
    public string Message { get; init; } = "";
    public string Ref { get; init; } = "";
}

public static class GitRunner
{
    public static string? FindWorktreeRoot(string? startPath)
    {
        try
        {
            var current = Path.GetFullPath(startPath ?? Directory.GetCurrentDirectory());
            if (File.Exists(current))
            {
                current = Path.GetDirectoryName(current) ?? current;
            }

            while (!string.IsNullOrEmpty(current))
            {
                var marker = Path.Combine(current, ".git");
                if (Directory.Exists(marker) || File.Exists(marker))
                {
                    return current;
                }

                var parent = Path.GetDirectoryName(current);
                if (parent == current)
                {
                    break;
                }

                current = parent ?? "";
            }

            var result = Run(startPath, ["rev-parse", "--show-toplevel"], 15);
            return result.ExitCode == 0 && !string.IsNullOrWhiteSpace(result.StdOut)
                ? Path.GetFullPath(result.StdOut.Trim())
                : null;
        }
        catch
        {
            return null;
        }
    }

    public static GitCommitInfo? LogInfo(string repoRoot, string gitRef, string relPath)
    {
        var result = Run(repoRoot, ["log", "-1", "--format=%H%n%h%n%an%n%ae%n%ci%n%s", gitRef, "--", relPath], 5);
        if (result.ExitCode != 0)
        {
            return null;
        }

        var parts = result.StdOut.Replace("\r", "").Trim().Split('\n');
        if (parts.Length < 6)
        {
            return null;
        }

        return new GitCommitInfo
        {
            Hash = parts[0],
            ShortHash = parts[1],
            Author = parts[2],
            Email = parts[3],
            Date = parts[4],
            Message = parts[5],
            Ref = gitRef,
        };
    }

    public static (GitCommitInfo? Local, GitCommitInfo? Remote) MergeInfo(string contextPath)
    {
        var repo = FindWorktreeRoot(contextPath);
        if (repo is null)
        {
            return (null, null);
        }

        var rel = Path.GetRelativePath(repo, Path.GetFullPath(contextPath)).Replace('\\', '/');
        if (rel.StartsWith("..", StringComparison.Ordinal))
        {
            return (null, null);
        }

        var local = LogInfo(repo, "HEAD", rel);
        GitCommitInfo? remote = null;
        foreach (var gitRef in new[] { "MERGE_HEAD", "REBASE_HEAD", "CHERRY_PICK_HEAD" })
        {
            remote = LogInfo(repo, gitRef, rel);
            if (remote is not null)
            {
                break;
            }
        }

        return (local, remote);
    }

    public static CommandResult Run(string? cwd, IEnumerable<string> args, int timeoutSeconds)
    {
        var start = new ProcessStartInfo("git")
        {
            WorkingDirectory = Directory.Exists(cwd) ? cwd : Path.GetDirectoryName(cwd) ?? Directory.GetCurrentDirectory(),
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8,
        };
        foreach (var arg in args)
        {
            start.ArgumentList.Add(arg);
        }

        try
        {
            using var process = Process.Start(start);
            if (process is null)
            {
                return new CommandResult(1, "", "无法启动 git");
            }

            if (!process.WaitForExit(timeoutSeconds * 1000))
            {
                try { process.Kill(true); } catch { }
                return new CommandResult(1, "", "git 超时");
            }

            return new CommandResult(process.ExitCode, process.StandardOutput.ReadToEnd(), process.StandardError.ReadToEnd());
        }
        catch (Exception ex)
        {
            return new CommandResult(1, "", ex.Message);
        }
    }

    public readonly record struct CommandResult(int ExitCode, string StdOut, string StdErr);
}
