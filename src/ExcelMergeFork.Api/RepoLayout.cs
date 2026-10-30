using ExcelMergeFork.Core.Excel;

namespace ExcelMergeFork.Api;

public static class RepoLayout
{
    public static string RepoRoot { get; private set; } = Directory.GetCurrentDirectory();

    public static string WebRoot => Path.Combine(RepoRoot, "web");

    public static string TestData => Path.Combine(RepoRoot, "TestData");

    public static string WorkRoot => Path.Combine(Path.GetTempPath(), "ExcelMergeForkWeb");

    public static void Locate(IEnumerable<string?> starts)
    {
        foreach (var start in starts)
        {
            if (string.IsNullOrWhiteSpace(start))
            {
                continue;
            }

            var full = Path.GetFullPath(start);
            if (File.Exists(Path.Combine(full, "web", "index.html")))
            {
                RepoRoot = full;
                return;
            }
        }

        foreach (var start in starts)
        {
            var found = Walk(start, dir => File.Exists(Path.Combine(dir, "TestData", "mode_a_local.xlsx")));
            if (found is not null)
            {
                RepoRoot = found;
                return;
            }
        }

        foreach (var start in starts)
        {
            var found = Walk(start, dir => File.Exists(Path.Combine(dir, "web", "index.html")));
            if (found is not null)
            {
                RepoRoot = found;
                return;
            }
        }
    }

    public static string ResolvePath(string raw)
    {
        var trimmed = raw.Trim().Trim('"');
        if (trimmed.Length == 0)
        {
            throw new InvalidOperationException("路径为空");
        }

        var path = Path.IsPathRooted(trimmed) ? trimmed : Path.Combine(RepoRoot, trimmed);
        return Path.GetFullPath(path);
    }

    public static string RequireWorkbook(string? raw, string role)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            throw new InvalidOperationException(role + " 路径为空");
        }

        var path = ResolvePath(raw);
        if (!ExcelFormats.MergeDiffSupported(path))
        {
            throw new InvalidOperationException(role + " 只支持 .xlsx / .xltx");
        }

        if (!File.Exists(path))
        {
            throw new FileNotFoundException(role + " 不存在: " + path, path);
        }

        return path;
    }

    public static string ResolveOutput(string? requested, string fallback, params string[] inputs)
    {
        var path = string.IsNullOrWhiteSpace(requested) ? Path.GetFullPath(fallback) : ResolvePath(requested);
        if (!ExcelFormats.MergeDiffSupported(path))
        {
            throw new InvalidOperationException("合并结果只支持 .xlsx / .xltx");
        }

        foreach (var input in inputs)
        {
            if (string.Equals(path, Path.GetFullPath(input), StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException("合并结果不能覆盖输入文件");
            }
        }

        var dir = Path.GetDirectoryName(path);
        if (string.IsNullOrEmpty(dir))
        {
            throw new InvalidOperationException("合并结果路径无效");
        }

        Directory.CreateDirectory(dir);
        return path;
    }

    public static bool SameOrigin(string origin, HostString host)
    {
        if (!Uri.TryCreate(origin, UriKind.Absolute, out var uri))
        {
            return false;
        }

        if (!string.Equals(uri.Host, host.Host, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return uri.Port == (host.Port ?? (uri.Scheme == "https" ? 443 : 80));
    }

    public static string SafeFileName(string? name)
    {
        var file = Path.GetFileName(string.IsNullOrWhiteSpace(name) ? "book.xlsx" : name);
        foreach (var c in Path.GetInvalidFileNameChars())
        {
            file = file.Replace(c, '_');
        }

        return file.Length == 0 ? "book.xlsx" : file;
    }

    private static string? Walk(string? start, Func<string, bool> match)
    {
        var dir = string.IsNullOrWhiteSpace(start) ? null : Path.GetFullPath(start);
        while (!string.IsNullOrEmpty(dir))
        {
            if (match(dir))
            {
                return dir;
            }

            dir = Path.GetDirectoryName(dir);
        }

        return null;
    }
}
