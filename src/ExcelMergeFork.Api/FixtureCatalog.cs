namespace ExcelMergeFork.Api;

public static class FixtureCatalog
{
    public static object Build()
    {
        var root = RepoLayout.TestData;
        var merges = new List<object>();
        var compares = new List<object>();
        AddMerge(merges, root, "mode-a", "模式 A", "A", "mode_a_local.xlsx", "mode_a_local.xlsx", "mode_a_remote.xlsx", "基准使用本地文件，与自动化测试一致");
        AddMerge(merges, root, "mode-b", "模式 B", "B", "mode_b_local.xlsx", "mode_b_local.xlsx", "mode_b_remote.xlsx", "基准使用本地文件，与自动化测试一致");
        AddMerge(merges, root, "mode-c", "模式 C", "C", "mode_c_local.xlsx", "mode_c_local.xlsx", "mode_c_remote.xlsx", "基准使用本地文件，与自动化测试一致");
        AddMerge(merges, root, "mode-d", "模式 D", "D", "mode_d_local.xlsx", "mode_d_local.xlsx", "mode_d_remote.xlsx", "基准使用本地文件，与自动化测试一致");
        AddMerge(merges, root, "mode-e", "模式 E", "E", "mode_e_local.xlsx", "mode_e_local.xlsx", "mode_e_remote.xlsx", "基准使用本地文件，与自动化测试一致");
        AddMerge(merges, root, "sample", "三文件样例", null, "local.xlsx", "base.xlsx", "remote.xlsx", "Local、Base、Remote 各用一份");
        AddMerge(merges, root, "language", "语言表", null, "Y语言表_Data_Language_local.xlsx", "Y语言表_Data_Language_base.xlsx", "Y语言表_Data_Language_remote.xlsx", "完整三方语言表");

        AddCompare(compares, root, "sample", "样例对比", "remote.xlsx", "local.xlsx");
        AddCompare(compares, root, "mode-d", "模式 D 对比", "mode_d_remote.xlsx", "mode_d_local.xlsx");
        AddCompare(compares, root, "language", "语言表对比", "Y语言表_Data_Language_remote.xlsx", "Y语言表_Data_Language_local.xlsx");

        var files = Directory.Exists(root)
            ? Directory.EnumerateFiles(root, "*.xlsx", SearchOption.TopDirectoryOnly)
                .Select(Path.GetFileName)
                .Where(name => name is not null)
                .Select(name => name!)
                .Where(name => !name.StartsWith("~$", StringComparison.Ordinal))
                .Where(name => !name.EndsWith("_compare.xlsx", StringComparison.OrdinalIgnoreCase))
                .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                .Select(name => new { Name = name, Path = Path.Combine(root, name) })
                .ToList()
            : [];

        return new
        {
            TestData = root,
            Merges = merges,
            Compares = compares,
            Files = files,
        };
    }

    private static void AddMerge(
        List<object> target,
        string root,
        string id,
        string title,
        string? mode,
        string local,
        string baseName,
        string remote,
        string note)
    {
        var localPath = Path.Combine(root, local);
        var basePath = Path.Combine(root, baseName);
        var remotePath = Path.Combine(root, remote);
        if (!File.Exists(localPath) || !File.Exists(basePath) || !File.Exists(remotePath))
        {
            return;
        }

        target.Add(new
        {
            Id = id,
            Title = title,
            Mode = mode,
            Note = note,
            Local = localPath,
            Base = basePath,
            Remote = remotePath,
        });
    }

    private static void AddCompare(List<object> target, string root, string id, string title, string left, string right)
    {
        var leftPath = Path.Combine(root, left);
        var rightPath = Path.Combine(root, right);
        if (!File.Exists(leftPath) || !File.Exists(rightPath))
        {
            return;
        }

        target.Add(new
        {
            Id = id,
            Title = title,
            Left = leftPath,
            Right = rightPath,
            Note = "左边是线上，右边是本地",
        });
    }
}
