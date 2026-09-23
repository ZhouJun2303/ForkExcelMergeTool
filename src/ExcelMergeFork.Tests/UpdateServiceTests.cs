using ExcelMergeFork.Core.Update;

namespace ExcelMergeFork.Tests;

public class UpdateServiceTests
{
    private const string ReleaseJson = """
        {
          "tag_name": "v3.2",
          "assets": [
            {
              "name": "ExcelMergeFork.exe",
              "browser_download_url": "https://example.test/ExcelMergeFork.exe"
            },
            {
              "name": "ExcelMergeFork.exe.sha256",
              "browser_download_url": "https://example.test/ExcelMergeFork.exe.sha256"
            },
            {
              "name": "ExcelMergeFork-package.zip",
              "browser_download_url": "https://example.test/ExcelMergeFork-package.zip"
            }
          ]
        }
        """;

    [Fact]
    public void ParseRelease_V30SeesV32AsNewer()
    {
        var info = UpdateService.ParseRelease(ReleaseJson, "3.0");
        Assert.NotNull(info);
        Assert.Equal("v3.2", info.Tag);
        Assert.Equal("3.2", info.Version);
        Assert.Equal("https://example.test/ExcelMergeFork.exe", info.DownloadUrl);
        Assert.Equal("https://example.test/ExcelMergeFork.exe.sha256", info.Sha256Url);
        Assert.True(info.IsNewer);
    }

    [Fact]
    public void ParseRelease_SameVersionIsNotNewer()
    {
        var info = UpdateService.ParseRelease(ReleaseJson, "3.2");
        Assert.NotNull(info);
        Assert.False(info.IsNewer);
    }

    [Fact]
    public void ParseRelease_MissingExeReturnsNull()
    {
        const string json = """
            { "tag_name": "v3.2", "assets": [ { "name": "notes.txt", "browser_download_url": "https://example.test/notes.txt" } ] }
            """;
        Assert.Null(UpdateService.ParseRelease(json, "3.0"));
    }

    [Theory]
    [InlineData("3.2", "3.0", true)]
    [InlineData("v3.2", "3.0", true)]
    [InlineData("3.0", "3.2", false)]
    [InlineData("3.2", "3.2", false)]
    [InlineData("3.10", "3.2", true)]
    [InlineData("3.2.1", "3.2", true)]
    public void IsNewer_ComparesNumericParts(string remote, string local, bool newer)
    {
        Assert.Equal(newer, UpdateService.IsNewer(remote, local));
    }

    [Fact]
    public void ReadExpectedSha256_ReadsHashFromReleaseFile()
    {
        var hash = new string('a', 64);
        Assert.Equal(hash, UpdateService.ReadExpectedSha256($"{hash}  ExcelMergeFork.exe"));
        Assert.Null(UpdateService.ReadExpectedSha256("not a hash"));
    }

    [Fact]
    public void EnsureSha256_RejectsMismatch()
    {
        var dir = Path.Combine(Path.GetTempPath(), "emf-sha-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var file = Path.Combine(dir, "payload.bin");
            File.WriteAllText(file, "excel-merge");
            var actual = UpdateService.Sha256File(file);
            UpdateService.EnsureSha256(file, actual.ToUpperInvariant());
            var wrong = new string('b', 64);
            var ex = Assert.Throws<UpdateException>(() => UpdateService.EnsureSha256(file, wrong));
            Assert.Contains("sha256", ex.Message);
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }

    [Fact]
    public void BuildApplyScript_QuotesPathsAndRetriesWhileLocked()
    {
        var pending = new PendingUpdate
        {
            ScriptPath = "unused",
            NewExe = @"C:\updates\a'b\ExcelMergeFork.exe",
            TargetExe = @"D:\工具\ExcelMergeFork.exe",
            BackupExe = @"D:\工具\ExcelMergeFork.exe.bak",
            LogPath = @"D:\工具\MergeExcelFork.update.log",
        };
        var script = UpdateService.BuildApplyScript(pending, initialWaitSeconds: 0, retryCount: 5);
        Assert.Contains(@"'C:\updates\a''b\ExcelMergeFork.exe'", script);
        Assert.Contains(@"'D:\工具\ExcelMergeFork.exe'", script);
        Assert.Contains("Copy-Item -LiteralPath $new -Destination $target", script);
        Assert.Contains("for ($i = 0; $i -lt 5; $i++)", script);
    }

    [Fact]
    public async Task ApplyScript_ReplacesTargetAfterItUnlocks()
    {
        var dir = Path.Combine(Path.GetTempPath(), "emf-upd-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        try
        {
            var target = Path.Combine(dir, "ExcelMergeFork.exe");
            var newerDir = Path.Combine(dir, "new");
            Directory.CreateDirectory(newerDir);
            var newer = Path.Combine(newerDir, "ExcelMergeFork.exe");
            await File.WriteAllTextAsync(target, "old");
            await File.WriteAllTextAsync(newer, "new-bytes");
            var pending = UpdateService.PrepareApply(newer, target, initialWaitSeconds: 0, retryCount: 12);
            using (new FileStream(target, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                UpdateService.LaunchApplyScript(pending);
                await Task.Delay(1500);
            }

            var deadline = DateTime.UtcNow.AddSeconds(15);
            string? content = null;
            while (DateTime.UtcNow < deadline)
            {
                try
                {
                    content = await File.ReadAllTextAsync(target);
                    if (content == "new-bytes")
                    {
                        break;
                    }
                }
                catch (IOException)
                {
                }

                await Task.Delay(300);
            }

            Assert.Equal("new-bytes", content);
            Assert.Contains("update ok", await File.ReadAllTextAsync(pending.LogPath));
        }
        finally
        {
            try
            {
                Directory.Delete(dir, true);
            }
            catch (IOException)
            {
            }
        }
    }
}
