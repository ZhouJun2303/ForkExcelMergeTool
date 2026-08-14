using System.Text.Json.Nodes;
using ExcelMergeFork.Core.Excel;

namespace ExcelMergeFork.Core.Git;

public sealed class IntegrationStatus
{
    public bool Installed { get; init; }
    public string Detail { get; init; } = "";
    public string Path { get; init; } = "";
}

public static class ForkIntegration
{
    public static string SettingsPath =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Fork", "settings.json");

    public static IntegrationStatus Status(string? toolPath = null)
    {
        try
        {
            if (!File.Exists(SettingsPath))
            {
                return new IntegrationStatus { Detail = "未找到 Fork settings.json" };
            }

            var json = JsonNode.Parse(File.ReadAllText(SettingsPath)) as JsonObject;
            var merge = json?["MergeTool"]?["Path"]?.GetValue<string>() ?? "";
            var current = Path.GetFullPath(toolPath ?? Environment.ProcessPath ?? "");
            var installed = !string.IsNullOrWhiteSpace(merge) &&
                            string.Equals(Path.GetFullPath(merge), current, StringComparison.OrdinalIgnoreCase);
            return new IntegrationStatus
            {
                Installed = installed,
                Path = SettingsPath,
                Detail = installed ? "Fork 已指向本工具" : "Fork 尚未指向本工具",
            };
        }
        catch (Exception ex)
        {
            return new IntegrationStatus { Detail = ex.Message };
        }
    }

    public static IntegrationStatus Install(string? toolPath = null)
    {
        var exe = Path.GetFullPath(toolPath ?? CurrentExecutable());
        Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
        var json = File.Exists(SettingsPath)
            ? JsonNode.Parse(File.ReadAllText(SettingsPath)) as JsonObject ?? new JsonObject()
            : new JsonObject();
        json["MergeTool"] = new JsonObject
        {
            ["Type"] = "Custom",
            ["Path"] = exe,
            ["Arguments"] = "$LOCAL,$BASE,$REMOTE,$MERGED",
        };
        json["DiffTool"] = new JsonObject
        {
            ["Type"] = "Custom",
            ["Path"] = exe,
            ["Arguments"] = "\"$REMOTE\" \"$LOCAL\"",
        };
        File.WriteAllText(SettingsPath, json.ToJsonString(new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));
        return Status(exe);
    }

    public static IntegrationStatus Uninstall(string? toolPath = null)
    {
        if (!File.Exists(SettingsPath))
        {
            return Status(toolPath);
        }

        var json = JsonNode.Parse(File.ReadAllText(SettingsPath)) as JsonObject;
        json?.Remove("MergeTool");
        json?.Remove("DiffTool");
        if (json is not null)
        {
            File.WriteAllText(SettingsPath, json.ToJsonString(new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));
        }

        return Status(toolPath);
    }

    public static string CurrentExecutable() =>
        Environment.ProcessPath ?? Path.Combine(AppPaths.Home, "ExcelMergeFork.exe");
}

public static class GitIntegration
{
    private static readonly string[] LegacyAttrLines =
    [
        "*.xlsx merge=excelmergefork",
        "*.XLSX merge=excelmergefork",
    ];

    public static string ToGitPath(string path) => Path.GetFullPath(path).Replace('\\', '/');

    public static string DriverCommand(string? toolPath = null)
    {
        var exe = ToGitPath(toolPath ?? ForkIntegration.CurrentExecutable());
        return $"\"{exe}\" --git-merge-driver \"%O\" \"%A\" \"%B\" \"%P\"";
    }

    public static string ResolveAttributesFilePath(string? configured, string userProfile)
    {
        if (!string.IsNullOrWhiteSpace(configured))
        {
            var path = Environment.ExpandEnvironmentVariables(configured.Trim());
            if (path.StartsWith("~", StringComparison.Ordinal))
            {
                path = Path.Combine(userProfile, path[1..].TrimStart('/', '\\'));
            }

            return path;
        }

        return Path.Combine(userProfile, ".config", "git", "attributes");
    }

    public static string AttributesFilePath()
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var result = GitRunner.Run(home, ["config", "--global", "--get", "core.attributesFile"], 15);
        var configured = result.ExitCode == 0 ? result.StdOut : null;
        return ResolveAttributesFilePath(configured, home);
    }

    public static IntegrationStatus Status()
    {
        var attributes = AttributesFilePath();
        var attrLines = ReadLines(attributes);
        var expected = ExcelFormats.GitAttrLines();
        var attrsInstalled = expected.All(attrLines.Contains);
        var driver = GitRunner.Run(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ["config", "--global", "--get", "merge.excelmergefork.driver"], 15);
        var installed = driver.ExitCode == 0 && attrsInstalled;
        return new IntegrationStatus
        {
            Installed = installed,
            Path = attributes,
            Detail = installed ? "已写入用户级 Git merge driver" : "尚未安装全局 Git 注入",
        };
    }

    public static IntegrationStatus Install(string? toolPath = null)
    {
        var command = DriverCommand(toolPath);
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        GitRunner.Run(home, ["config", "--global", "merge.excelmergefork.name", "ExcelMergeFork workbook merge driver"], 15);
        GitRunner.Run(home, ["config", "--global", "merge.excelmergefork.driver", command], 15);
        GitRunner.Run(home, ["config", "--global", "merge.excelmergefork.recursive", "binary"], 15);
        var attributes = AttributesFilePath();
        var existing = ReadLines(attributes);
        existing.RemoveAll(line => LegacyAttrLines.Contains(line));
        foreach (var line in ExcelFormats.GitAttrLines())
        {
            if (!existing.Contains(line))
            {
                existing.Add(line);
            }
        }

        WriteLines(attributes, existing);
        return Status();
    }

    public static IntegrationStatus Uninstall()
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        foreach (var key in new[] { "name", "driver", "recursive" })
        {
            GitRunner.Run(home, ["config", "--global", "--unset-all", $"merge.excelmergefork.{key}"], 15);
        }

        var attributes = AttributesFilePath();
        if (File.Exists(attributes))
        {
            var remove = ExcelFormats.GitAttrLines().Concat(LegacyAttrLines).ToHashSet(StringComparer.Ordinal);
            WriteLines(attributes, ReadLines(attributes).Where(line => !remove.Contains(line)).ToList());
        }

        return Status();
    }

    private static List<string> ReadLines(string path)
    {
        if (!File.Exists(path))
        {
            return [];
        }

        return File.ReadAllLines(path).ToList();
    }

    private static void WriteLines(string path, IReadOnlyList<string> lines)
    {
        var parent = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(parent))
        {
            Directory.CreateDirectory(parent);
        }

        var text = string.Join("\n", lines);
        if (lines.Count > 0)
        {
            text += "\n";
        }

        File.WriteAllText(path, text);
    }
}
