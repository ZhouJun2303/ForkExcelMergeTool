using ExcelMergeFork.Core.Git;

namespace ExcelMergeFork.Tests;

public class GitIntegrationTests
{
    [Fact]
    public void DriverCommand_QuotesGitPlaceholders()
    {
        var exe = Path.Combine(Path.GetTempPath(), "Excel Merge", "ExcelMergeFork.exe");
        var command = GitIntegration.DriverCommand(exe);
        Assert.Contains(" --git-merge-driver ", command);
        Assert.Contains("\"%O\"", command);
        Assert.Contains("\"%A\"", command);
        Assert.Contains("\"%B\"", command);
        Assert.Contains("\"%P\"", command);
        Assert.DoesNotContain(" %O ", command);
        Assert.StartsWith("\"", command);
    }

    [Fact]
    public void AttributesFilePath_UsesCoreAttributesFileOrXdg()
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var git = GitRunner.Run(home, ["config", "--global", "--get", "core.attributesFile"], 15);
        var configured = git.ExitCode == 0 ? git.StdOut : null;
        var expected = GitIntegration.ResolveAttributesFilePath(configured, home);
        Assert.Equal(expected, GitIntegration.AttributesFilePath());
        Assert.False(
            expected.EndsWith(Path.Combine(home, ".gitattributes"), StringComparison.OrdinalIgnoreCase) &&
            string.IsNullOrWhiteSpace(configured),
            "default attributes path must not be ~/.gitattributes");
        if (string.IsNullOrWhiteSpace(configured))
        {
            Assert.Equal(Path.Combine(home, ".config", "git", "attributes"), expected);
        }
    }

    [Fact]
    public void ResolveAttributesFilePath_PrefersConfiguredThenXdg()
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        Assert.Equal(
            Path.Combine(home, ".config", "git", "attributes"),
            GitIntegration.ResolveAttributesFilePath("  ", home));
        Assert.Equal(@"D:\custom\attrs", GitIntegration.ResolveAttributesFilePath(@"D:\custom\attrs", home));
    }

    [Fact]
    public void WindowCloseExitCode_IsZeroOnlyAfterWriteBack()
    {
        Assert.Equal(0, GitMergeDriver.WindowCloseExitCode(true));
        Assert.Equal(1, GitMergeDriver.WindowCloseExitCode(false));
    }

    [Fact]
    public void Startup_SetsExplicitShutdownBeforeLaunch()
    {
        var path = Path.Combine(TestRepo.Root, "src", "ExcelMergeFork.App", "App.xaml.cs");
        var text = File.ReadAllText(path);
        var startupAt = text.IndexOf("void OnStartup", StringComparison.Ordinal);
        Assert.True(startupAt >= 0, "missing OnStartup");
        var startup = text[startupAt..];
        var explicitAt = startup.IndexOf("ShutdownMode = ShutdownMode.OnExplicitShutdown", StringComparison.Ordinal);
        var activateAt = startup.IndexOf("DesktopLaunch.Activate", StringComparison.Ordinal);
        Assert.True(explicitAt >= 0, "OnStartup must set OnExplicitShutdown");
        Assert.True(activateAt > explicitAt, "OnExplicitShutdown must be set before the shared UI opens");

        var exitAt = text.IndexOf("int ExitCodeOf", StringComparison.Ordinal);
        Assert.True(exitAt >= 0, "missing ExitCodeOf");
        var exit = text[exitAt..];
        Assert.Contains("WindowCloseExitCode(_desktop?.WriteBackSucceeded ?? false)", exit, StringComparison.Ordinal);
    }
}
