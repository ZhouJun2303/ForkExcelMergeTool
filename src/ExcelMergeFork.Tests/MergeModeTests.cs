using ClosedXML.Excel;
using ExcelMergeFork.Core.Excel;
using ExcelMergeFork.Core.Merge;

namespace ExcelMergeFork.Tests;

public class MergeModeTests
{
    private static readonly MergeChoice[] PythonDeChoices =
    [
        new() { Sheet = "Data", Key = "k1", Choice = "local" },
        new() { Sheet = "Data", Key = "k2", Choice = "remote" },
    ];

    [Fact]
    public void ModeA_InsertsPrefixedRowsFromRemote()
    {
        var output = RunMode("A", "mode_a_local.xlsx", "mode_a_remote.xlsx");
        var keys = DataKeys(output, "Data");
        Assert.Equal(["A-1", "A-2", "B-1", "B-2"], keys);
        Assert.Equal(["A-2", "a2", "x"], Row(output, "Data", "A-2"));
        Assert.Equal(["B-2", "b2", "y"], Row(output, "Data", "B-2"));
        Assert.Equal(["A-1", "a1", "x"], Row(output, "Data", "A-1"));
    }

    [Fact]
    public void ModeB_InsertsExtraColumnsWithRemoteValues()
    {
        var output = RunMode("B", "mode_b_local.xlsx", "mode_b_remote.xlsx");
        using var wb = new XLWorkbook(output);
        var snap = SheetSnapshot.From(wb.Worksheet("Data"));
        Assert.Equal(["Key", "名称", "数量", "单价", "总价"], snap.Headers);
        var apple = snap.RowsByKey["1"];
        Assert.Equal("苹果", CellText.From(apple[1]));
        Assert.Equal("2.5", CellText.From(apple[3]));
        Assert.Equal("25", CellText.From(apple[4]));
    }

    [Fact]
    public void ModeC_AppendsSheetsOnlyOnOtherSide()
    {
        var output = RunMode("C", "mode_c_local.xlsx", "mode_c_remote.xlsx");
        using var wb = new XLWorkbook(output);
        var names = wb.Worksheets.Select(ws => ws.Name).ToList();
        Assert.Contains("Sheet1", names);
        Assert.Contains("Sheet2", names);
        Assert.Contains("Sheet3", names);
        var sheet2 = SheetSnapshot.From(wb.Worksheet("Sheet2"));
        Assert.Equal("x", CellText.From(sheet2.Rows[1][0]));
        Assert.Equal("1", CellText.From(sheet2.Rows[1][1]));
    }

    [Fact]
    public void ModeD_AppliesLocalAndRemoteRowChoices()
    {
        var output = RunMode("D", "mode_d_local.xlsx", "mode_d_remote.xlsx", PythonDeChoices);
        Assert.Equal(["k1", "local1", "local2"], Row(output, "Data", "k1"));
        Assert.Equal(["k2", "remote_a", "remote_b"], Row(output, "Data", "k2"));
    }

    [Fact]
    public void ModeE_CombinesNewRowNewSheetAndChoices()
    {
        var output = RunMode("E", "mode_e_local.xlsx", "mode_e_remote.xlsx", PythonDeChoices);
        using var wb = new XLWorkbook(output);
        var names = wb.Worksheets.Select(ws => ws.Name).ToList();
        Assert.Contains("Main", names);
        Assert.Contains("Extra", names);
        Assert.Equal(["A-1", "A-2", "id1"], DataKeys(output, "Main"));
        Assert.Equal(["A-2", "c", "d"], Row(output, "Main", "A-2"));
        Assert.Equal(["A-1", "a", "b"], Row(output, "Main", "A-1"));
        var extra = SheetSnapshot.From(wb.Worksheet("Extra"));
        Assert.Equal("1", CellText.From(extra.Rows[1][0]));
        Assert.Equal("2", CellText.From(extra.Rows[1][1]));
    }

    private static string RunMode(string mode, string localName, string remoteName, IReadOnlyList<MergeChoice>? choices = null)
    {
        var local = TestRepo.Fixture(localName);
        var remote = TestRepo.Fixture(remoteName);
        Assert.True(File.Exists(local), "缺少夹具 " + local);
        Assert.True(File.Exists(remote), "缺少夹具 " + remote);
        var merged = Path.Combine(Path.GetTempPath(), "emf-mode-" + mode + "-" + Guid.NewGuid().ToString("N") + ".xlsx");
        var result = MergeService.Run(
            local,
            local,
            remote,
            merged,
            mode: mode,
            choices: choices,
            createBackup: false);
        Assert.True(result.ExitCode == 0, result.Error ?? "merge failed");
        Assert.True(File.Exists(merged));
        return merged;
    }

    private static List<string> DataKeys(string path, string sheet)
    {
        using var wb = new XLWorkbook(path);
        var snap = SheetSnapshot.From(wb.Worksheet(sheet));
        return snap.OrderedKeys.Where(k => !string.Equals(k, "Key", StringComparison.OrdinalIgnoreCase)).ToList();
    }

    private static List<string> Row(string path, string sheet, string key)
    {
        using var wb = new XLWorkbook(path);
        var snap = SheetSnapshot.From(wb.Worksheet(sheet));
        Assert.True(snap.RowsByKey.ContainsKey(key), sheet + " 缺少 key " + key);
        return snap.RowsByKey[key].Select(CellText.From).ToList();
    }
}
