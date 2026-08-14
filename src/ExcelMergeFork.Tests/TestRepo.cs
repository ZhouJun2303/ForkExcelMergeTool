namespace ExcelMergeFork.Tests;

internal static class TestRepo
{
    public static string Root
    {
        get
        {
            var dir = AppContext.BaseDirectory;
            while (!string.IsNullOrEmpty(dir))
            {
                if (File.Exists(Path.Combine(dir, "TestData", "mode_a_local.xlsx")))
                {
                    return dir;
                }

                dir = Path.GetDirectoryName(dir) ?? "";
            }

            throw new DirectoryNotFoundException("找不到带 TestData/mode_a_local.xlsx 的仓库根目录");
        }
    }

    public static string TestData => Path.Combine(Root, "TestData");

    public static string Fixture(string name) => Path.Combine(TestData, name);
}
