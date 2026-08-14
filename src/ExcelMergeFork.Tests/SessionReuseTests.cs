using ExcelMergeFork.Core.Excel;
using ExcelMergeFork.Core.Merge;

namespace ExcelMergeFork.Tests;

public class SessionReuseTests
{
    [Fact]
    public void PreviewBuilder_ReadsAlreadyOpenSession_NotDiskPerRow()
    {
        var dir = Path.Combine(Path.GetTempPath(), "emf-session-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var local = Path.Combine(dir, "local.xlsx");
        var remote = Path.Combine(dir, "remote.xlsx");
        var baseline = Path.Combine(dir, "base.xlsx");
        File.Copy(TestRepo.Fixture("mode_a_local.xlsx"), local);
        File.Copy(TestRepo.Fixture("mode_a_local.xlsx"), baseline);
        File.Copy(TestRepo.Fixture("mode_a_remote.xlsx"), remote);

        using var session = new MergeSession(local, baseline, remote);
        File.Copy(TestRepo.Fixture("mode_a_local.xlsx"), remote, overwrite: true);

        var preview = PreviewBuilder.Build(session, MergeOptions.ForMode("A"));
        var added = preview.Items.Single(i => i.Key == "A-2");
        Assert.Equal("将新增行", added.Action);
        var remoteRow = session.Remote.Snapshot("Data").RowsByKey["A-2"];
        Assert.Equal(remoteRow.Select(CellText.From), added.RemoteValues);
    }
}
