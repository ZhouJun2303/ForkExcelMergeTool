using System.Globalization;
using ExcelMergeFork.Core.Excel;

namespace ExcelMergeFork.Core.Routing;

public enum LaunchMode
{
    Settings,
    Merge,
    Compare,
    GitDriver,
    InstallFork,
    UninstallFork,
    InstallGit,
    UninstallGit,
    Invalid,
}

public sealed class LaunchRequest
{
    public LaunchMode Mode { get; init; }
    public IReadOnlyList<string> Files { get; init; } = [];
    public bool IncludeSame { get; init; }

    public string? Local => Files.ElementAtOrDefault(Mode == LaunchMode.Compare ? 1 : 0);
    public string? Base => Files.ElementAtOrDefault(1);
    public string? Remote => Files.ElementAtOrDefault(Mode == LaunchMode.Compare ? 0 : 2);
    public string? Merged => Files.ElementAtOrDefault(3);
}

public static class LaunchArgs
{
    public static LaunchRequest Parse(string[] raw)
    {
        var args = raw.ToList();
        var includeSame = args.Remove("--include-same");
        if (args.Count == 0)
        {
            return new LaunchRequest { Mode = LaunchMode.Settings };
        }

        return args[0] switch
        {
            "--main" => new LaunchRequest { Mode = LaunchMode.Settings },
            "--install-fork-integration" => new LaunchRequest { Mode = LaunchMode.InstallFork, Files = args.Skip(1).ToList() },
            "--uninstall-fork-integration" => new LaunchRequest { Mode = LaunchMode.UninstallFork, Files = args.Skip(1).ToList() },
            "--install-git-integration" => new LaunchRequest { Mode = LaunchMode.InstallGit, Files = args.Skip(1).ToList() },
            "--uninstall-git-integration" => new LaunchRequest { Mode = LaunchMode.UninstallGit, Files = args.Skip(1).ToList() },
            "--git-merge-driver" => new LaunchRequest { Mode = LaunchMode.GitDriver, Files = args.Skip(1).ToList() },
            _ => ParseFiles(args, includeSame),
        };
    }

    public static IReadOnlyList<string> UnsupportedMergeDiff(IEnumerable<string?> paths) =>
        paths.Where(p => !string.IsNullOrWhiteSpace(p) && !ExcelFormats.MergeDiffSupported(p!)).Cast<string>().ToList();

    private static LaunchRequest ParseFiles(List<string> args, bool includeSame)
    {
        var files = new List<string>();
        if (args.Count is 2 or 4)
        {
            files.AddRange(args.Select(Clean));
        }
        else if (args.Count == 1)
        {
            files.AddRange(SplitCsv(args[0]));
        }
        else
        {
            foreach (var arg in args)
            {
                files.AddRange(arg.Contains(',') ? SplitCsv(arg) : [Clean(arg)]);
            }
        }

        files = files.Where(f => f.Length > 0).ToList();
        if (files.Count == 4)
        {
            return new LaunchRequest { Mode = LaunchMode.Merge, Files = files, IncludeSame = includeSame };
        }

        if (files.Count == 2)
        {
            return new LaunchRequest { Mode = LaunchMode.Compare, Files = files, IncludeSame = includeSame };
        }

        return new LaunchRequest { Mode = LaunchMode.Invalid, Files = files, IncludeSame = includeSame };
    }

    private static IEnumerable<string> SplitCsv(string value)
    {
        return value.Split(',').Select(Clean).Where(s => s.Length > 0);
    }

    private static string Clean(string value) => value.Trim().Trim('"').Trim('\'');
}
