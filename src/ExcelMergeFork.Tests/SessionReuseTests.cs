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
        Assert.Empty(added.LocalValues);
        Assert.Empty(added.BaseValues);
        Assert.Contains(added.RemoteValues, line => line.Contains("a2"));
        Assert.Contains(added.RemoteValues, line => line.StartsWith("Key:", StringComparison.Ordinal) || line.Contains("A-2"));
    }

    [Fact]
    public void SideTexts_MissingRow_IsEmpty_PresentRow_HasHeaders()
    {
        var local = TestRepo.Fixture("mode_a_local.xlsx");
        var remote = TestRepo.Fixture("mode_a_remote.xlsx");
        using var session = new MergeSession(local, local, remote);
        Assert.Empty(PreviewBuilder.SideTexts(session.Local, "Data", "A-2", column: false));
        Assert.Empty(PreviewBuilder.SideTexts(session.Base, "Data", "A-2", column: false));
        var remoteLines = PreviewBuilder.SideTexts(session.Remote, "Data", "A-2", column: false);
        Assert.NotEmpty(remoteLines);
        Assert.Contains(remoteLines, line => line.Contains("a2"));
    }
}
