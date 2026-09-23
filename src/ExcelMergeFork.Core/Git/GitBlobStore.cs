using System.Diagnostics;
using System.Text.RegularExpressions;

namespace ExcelMergeFork.Core.Git;

public static partial class GitBlobStore
{
    public static void WriteBlob(string repoRoot, string objectId, string destination)
    {
        if (string.IsNullOrWhiteSpace(objectId) || !ObjectIdPattern().IsMatch(objectId))
        {
            throw new ArgumentException("无效的 Git 对象：" + objectId);
        }

        var directory = Path.GetDirectoryName(destination);
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var start = new ProcessStartInfo("git")
        {
            WorkingDirectory = Directory.Exists(repoRoot) ? repoRoot : Directory.GetCurrentDirectory(),
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true,
        };
        start.ArgumentList.Add("cat-file");
        start.ArgumentList.Add("blob");
        start.ArgumentList.Add(objectId);

        using var process = Process.Start(start) ?? throw new InvalidOperationException("无法启动 git");
        using var file = new FileStream(destination, FileMode.Create, FileAccess.Write, FileShare.Read);
        var stderr = Task.Run(process.StandardError.ReadToEnd);
        process.StandardOutput.BaseStream.CopyTo(file);
        file.Flush();
        if (!process.WaitForExit(30_000))
        {
            try { process.Kill(true); } catch { }
            throw new TimeoutException("git cat-file 超时");
        }

        var error = stderr.GetAwaiter().GetResult();
        if (process.ExitCode != 0)
        {
            try { File.Delete(destination); } catch { }
            throw new InvalidOperationException(string.IsNullOrWhiteSpace(error) ? "git cat-file 失败" : error.Trim());
        }
    }

    [GeneratedRegex("^[0-9a-fA-F]{4,64}$")]
    private static partial Regex ObjectIdPattern();
}
