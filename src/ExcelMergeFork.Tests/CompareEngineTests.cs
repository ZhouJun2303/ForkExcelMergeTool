using ClosedXML.Excel;
using ExcelMergeFork.Core;
using ExcelMergeFork.Core.Compare;

namespace ExcelMergeFork.Tests;

public class CompareEngineTests
{
    [Fact]
    public void Compute_ReturnsStatusesWithoutWriting_WriteAppliesSpecFills()
    {
        var dir = Path.Combine(Path.GetTempPath(), "emf-compare-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var left = Path.Combine(dir, "left.xlsx");
            var right = Path.Combine(dir, "right.xlsx");
            Write(left, [
                ["Key", "V", "OnlyA"],
                ["keep", "same", "a"],
                ["changed", "old", ""],
                ["gone", "x", ""],
            ]);
            Write(right, [
                ["Key", "V", "OnlyB"],
                ["keep", "same", "b"],
                ["changed", "new", ""],
                ["fresh", "y", ""],
            ]);

            var result = CompareEngine.Compute(left, right);
            Assert.False(File.Exists(result.OutputPath), "Compute 不得写出对比工作簿");
            var changed = Assert.Single(result.Rows, r => r.Status == "修改" && r.Key == "changed");
            var value = Assert.Single(changed.Cells, c => c.Header == "V");
            Assert.True(value.Changed);
            Assert.Equal("old", value.Left);
            Assert.Equal("new", value.Right);
            Assert.Contains(result.Rows, r => r.Status == "新增行" && r.Key == "fresh");
            Assert.Contains(result.Rows, r => r.Status == "删除行" && r.Key == "gone");
            Assert.Contains(result.Rows, r => r.Status == "新增列" && r.Right.Contains("OnlyB"));
            Assert.Contains(result.Rows, r => r.Status == "删除列" && r.Left.Contains("OnlyA"));

            CompareEngine.WriteExcel(result);
            Assert.True(File.Exists(result.OutputPath));

            using var wb = new XLWorkbook(result.OutputPath);
            var ws = wb.Worksheet("Data");
            AssertFill(ws, "fresh", AppConstants.ColorNew);
            AssertFill(ws, "gone", AppConstants.ColorConflict);
            AssertFill(ws, "changed", AppConstants.ColorChanged);
        }
        finally
        {
            try { Directory.Delete(dir, true); } catch { }
        }
    }

    [Fact]
    public void Compute_DuplicateKeys_KeepsEarlierChangeAndExtraCopy()
    {
        var dir = Path.Combine(Path.GetTempPath(), "emf-compare-dup-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var left = Path.Combine(dir, "left.xlsx");
            var right = Path.Combine(dir, "right.xlsx");
            Write(left, [
                ["Key", "V"],
                ["dup", "old"],
                ["dup", "new"],
            ]);
            Write(right, [
                ["Key", "V"],
                ["dup", "new"],
            ]);

            var result = CompareEngine.Compute(left, right);
            var modified = Assert.Single(result.Rows, r => r.Status == "修改");
            Assert.Equal("dup", modified.Key);
            Assert.Equal(2, modified.LeftRow);
            Assert.Equal(2, modified.RightRow);
            Assert.Contains(modified.Cells, c => c.Changed && c.Header == "V" && c.Left == "old" && c.Right == "new");

            var extra = Assert.Single(result.Rows, r => r.Status == "删除行");
            Assert.Equal("dup #2", extra.Key);
            Assert.Equal(3, extra.LeftRow);
            Assert.Equal(0, extra.RightRow);
        }
        finally
        {
            try { Directory.Delete(dir, true); } catch { }
        }
    }

    [Fact]
    public void Compute_IdenticalDuplicateKeys_AreNotDiffs()
    {
        var dir = Path.Combine(Path.GetTempPath(), "emf-compare-same-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var left = Path.Combine(dir, "left.xlsx");
            var right = Path.Combine(dir, "right.xlsx");
            var rows = new[]
            {
                new[] { "Key", "V" },
                new[] { "dup", "one" },
                new[] { "dup", "two" },
            };
            Write(left, rows);
            Write(right, rows);

            var result = CompareEngine.Compute(left, right);
            Assert.Empty(result.Rows);
        }
        finally
        {
            try { Directory.Delete(dir, true); } catch { }
        }
    }

    [Fact]
    public void CompareAndWrite_MissingFile_ReturnsExit1()
    {
        Assert.Equal(1, CompareEngine.CompareAndWrite(
            Path.Combine(Path.GetTempPath(), "missing-a.xlsx"),
            Path.Combine(Path.GetTempPath(), "missing-b.xlsx")));
    }

    private static void AssertFill(IXLWorksheet ws, string key, string hex)
    {
        var last = ws.LastRowUsed()?.RowNumber() ?? 1;
        for (var r = 2; r <= last; r++)
        {
            if (ws.Cell(r, 1).GetString() != key)
            {
                continue;
            }

            var actual = ws.Cell(r, 4).Style.Fill.BackgroundColor;
            var expected = XLColor.FromHtml("#" + hex);
            Assert.True(
                actual.Color.ToArgb() == expected.Color.ToArgb(),
                $"key {key} fill {actual} != #{hex}");
            return;
        }

        Assert.Fail("对比表里没有 key " + key);
    }

    private static void Write(string path, IEnumerable<IEnumerable<object>> rows)
    {
        using var wb = new XLWorkbook();
        var ws = wb.AddWorksheet("Data");
        var r = 1;
        foreach (var row in rows)
        {
            var c = 1;
            foreach (var value in row)
            {
                ws.Cell(r, c).Value = value.ToString();
                c++;
            }

            r++;
        }

        wb.SaveAs(path);
    }
}
