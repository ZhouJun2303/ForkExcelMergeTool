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
    public void StartGitDriver_SetsExplicitShutdownBeforeShow()
    {
        var path = Path.Combine(TestRepo.Root, "src", "ExcelMergeFork.App", "App.xaml.cs");
        var text = File.ReadAllText(path);
        var start = text.IndexOf("private void StartGitDriver", StringComparison.Ordinal);
        Assert.True(start >= 0, "missing StartGitDriver");
        var end = text.IndexOf("private static string? ResolveFeature", start, StringComparison.Ordinal);
        Assert.True(end > start, "cannot isolate StartGitDriver");
        var method = text[start..end];
        Assert.Contains("ShowAndTrack(window", method, StringComparison.Ordinal);
        Assert.Contains("WindowCloseExitCode(window.WriteBackSucceeded)", method, StringComparison.Ordinal);

        var startup = text[text.IndexOf("private void OnStartup", StringComparison.Ordinal)..];
        var explicitAt = startup.IndexOf("ShutdownMode = ShutdownMode.OnExplicitShutdown", StringComparison.Ordinal);
        var resolveAt = startup.IndexOf("ResolveFeature(", StringComparison.Ordinal);
        Assert.True(explicitAt >= 0, "OnStartup must set OnExplicitShutdown");
        Assert.True(resolveAt > explicitAt, "OnExplicitShutdown must be set before any choice dialog");
    }
}
