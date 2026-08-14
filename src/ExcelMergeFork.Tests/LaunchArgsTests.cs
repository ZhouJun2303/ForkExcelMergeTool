using ExcelMergeFork.Core.Merge;
using ExcelMergeFork.Core.Routing;

namespace ExcelMergeFork.Tests;

public class LaunchArgsTests
{
    [Fact]
    public void NoArgs_OpensSettings()
    {
        Assert.Equal(LaunchMode.Settings, LaunchArgs.Parse([]).Mode);
    }

    [Fact]
    public void FourPaths_AreMerge()
    {
        var req = LaunchArgs.Parse(["a.xlsx", "b.xlsx", "c.xlsx", "d.xlsx"]);
        Assert.Equal(LaunchMode.Merge, req.Mode);
        Assert.Equal("a.xlsx", req.Local);
        Assert.Equal("b.xlsx", req.Base);
        Assert.Equal("c.xlsx", req.Remote);
        Assert.Equal("d.xlsx", req.Merged);
    }

    [Fact]
    public void TwoPaths_AreCompare_RemoteFirst()
    {
        var req = LaunchArgs.Parse(["remote.xlsx", "local.xlsx"]);
        Assert.Equal(LaunchMode.Compare, req.Mode);
        Assert.Equal("remote.xlsx", req.Remote);
        Assert.Equal("local.xlsx", req.Local);
    }

    [Fact]
    public void CommaJoinedFourPaths_AreMerge()
    {
        var req = LaunchArgs.Parse(["a.xlsx,b.xlsx,c.xlsx,d.xlsx"]);
        Assert.Equal(LaunchMode.Merge, req.Mode);
        Assert.Equal(4, req.Files.Count);
    }

    [Fact]
    public void GitDriverFlag_IsParsed()
    {
        var req = LaunchArgs.Parse(["--git-merge-driver", "o", "a", "b", "p"]);
        Assert.Equal(LaunchMode.GitDriver, req.Mode);
        Assert.Equal(["o", "a", "b", "p"], req.Files);
    }

    [Theory]
    [InlineData("--install-fork-integration", LaunchMode.InstallFork)]
    [InlineData("--uninstall-fork-integration", LaunchMode.UninstallFork)]
    [InlineData("--install-git-integration", LaunchMode.InstallGit)]
    [InlineData("--uninstall-git-integration", LaunchMode.UninstallGit)]
    [InlineData("--main", LaunchMode.Settings)]
    public void IntegrationFlags_MatchContract(string flag, LaunchMode mode)
    {
        Assert.Equal(mode, LaunchArgs.Parse([flag]).Mode);
    }

    [Fact]
    public void OddPathCount_IsInvalid()
    {
        Assert.Equal(LaunchMode.Invalid, LaunchArgs.Parse(["only.xlsx"]).Mode);
    }

    [Fact]
    public void MergeService_MissingInputs_ReturnsExit1()
    {
        var missing = Path.Combine(Path.GetTempPath(), "no-such-merge.xlsx");
        var result = MergeService.Run(missing, missing, missing, missing, mode: "A", createBackup: false);
        Assert.Equal(1, result.ExitCode);
    }
}
