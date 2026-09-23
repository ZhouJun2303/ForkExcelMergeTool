using System.Diagnostics;
using System.Net.Http;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace ExcelMergeFork.Core.Update;

public sealed class UpdateInfo
{
    public required string Tag { get; init; }
    public required string Version { get; init; }
    public required string DownloadUrl { get; init; }
    public string? Sha256Url { get; init; }
    public bool IsNewer { get; init; }
}

public readonly record struct UpdateProgress(string Status);

public sealed class PendingUpdate
{
    public required string ScriptPath { get; init; }
    public required string NewExe { get; init; }
    public required string TargetExe { get; init; }
    public required string BackupExe { get; init; }
    public required string LogPath { get; init; }
}

public sealed class UpdateException : Exception
{
    public UpdateException(string message) : base(message)
    {
    }

    public UpdateException(string message, Exception inner) : base(message, inner)
    {
    }
}

public static class UpdateService
{
    private static readonly Regex Sha256Pattern = new(@"\b[a-fA-F0-9]{64}\b", RegexOptions.Compiled);

    public static string? CurrentExecutable()
    {
        var path = Environment.ProcessPath;
        if (string.IsNullOrWhiteSpace(path))
        {
            return null;
        }

        if (!Path.GetFileName(path).Equals(AppConstants.UpdateAssetName, StringComparison.OrdinalIgnoreCase))
        {
            return null;
        }

        return Path.GetFullPath(path);
    }

    public static async Task<UpdateInfo?> CheckAsync(CancellationToken token = default)
    {
        AppLog.Info("正在检查更新...");
        try
        {
            using var client = CreateClient(TimeSpan.FromSeconds(15));
            client.DefaultRequestHeaders.TryAddWithoutValidation("Accept", "application/vnd.github+json");
            var url = $"https://api.github.com/repos/{AppConstants.GitHubRepo}/releases/latest";
            using var response = await client.GetAsync(url, token);
            if (!response.IsSuccessStatusCode)
            {
                throw new UpdateException($"检查更新失败: HTTP {(int)response.StatusCode}");
            }

            var json = await response.Content.ReadAsStringAsync(token);
            return ParseRelease(json, AppVersion.Display);
        }
        catch (UpdateException)
        {
            throw;
        }
        catch (Exception ex) when (ex is HttpRequestException or TaskCanceledException or JsonException)
        {
            AppLog.Exception("检查更新失败", ex);
            throw new UpdateException("检查更新失败: " + ex.Message, ex);
        }
    }

    public static UpdateInfo? ParseRelease(string json, string localVersion)
    {
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        var tag = root.TryGetProperty("tag_name", out var tagNode) ? tagNode.GetString() ?? "" : "";
        var version = tag.TrimStart('v', 'V');
        string? download = null;
        string? sha = null;
        if (root.TryGetProperty("assets", out var assets))
        {
            foreach (var asset in assets.EnumerateArray())
            {
                var name = asset.TryGetProperty("name", out var nameNode) ? nameNode.GetString() ?? "" : "";
                var href = asset.TryGetProperty("browser_download_url", out var hrefNode) ? hrefNode.GetString() : null;
                if (string.IsNullOrWhiteSpace(href))
                {
                    continue;
                }

                if (name.Equals(AppConstants.UpdateAssetName, StringComparison.OrdinalIgnoreCase))
                {
                    download = href;
                }
                else if (name.Equals(AppConstants.UpdateSha256AssetName, StringComparison.OrdinalIgnoreCase))
                {
                    sha = href;
                }
            }
        }

        if (string.IsNullOrWhiteSpace(download) || string.IsNullOrWhiteSpace(version))
        {
            return null;
        }

        return new UpdateInfo
        {
            Tag = string.IsNullOrWhiteSpace(tag) ? version : tag,
            Version = version,
            DownloadUrl = download,
            Sha256Url = sha,
            IsNewer = IsNewer(version, localVersion),
        };
    }

    public static bool IsNewer(string remote, string local)
    {
        static int[] Parts(string text) =>
            text.Trim().TrimStart('v', 'V')
                .Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Select(p => int.TryParse(new string(p.TakeWhile(char.IsDigit).ToArray()), out var n) ? n : 0)
                .ToArray();

        var a = Parts(remote);
        var b = Parts(local);
        var n = Math.Max(a.Length, b.Length);
        for (var i = 0; i < n; i++)
        {
            var left = i < a.Length ? a[i] : 0;
            var right = i < b.Length ? b[i] : 0;
            if (left != right)
            {
                return left > right;
            }
        }

        return false;
    }

    public static async Task<string> DownloadAsync(
        UpdateInfo info,
        IProgress<UpdateProgress>? progress = null,
        CancellationToken token = default)
    {
        if (string.IsNullOrWhiteSpace(info.DownloadUrl))
        {
            throw new UpdateException("最新 Release 未找到 " + AppConstants.UpdateAssetName);
        }

        var dir = Path.Combine(Path.GetTempPath(), "ExcelMergeForkUpdate_" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var exePath = Path.Combine(dir, AppConstants.UpdateAssetName);
        try
        {
            await DownloadFileAsync(info.DownloadUrl, exePath, info.Version, progress, token);
            if (!string.IsNullOrWhiteSpace(info.Sha256Url))
            {
                progress?.Report(new UpdateProgress("正在校验..."));
                var shaPath = Path.Combine(dir, AppConstants.UpdateSha256AssetName);
                await DownloadFileAsync(info.Sha256Url, shaPath, info.Version, progress: null, token);
                var expected = ReadExpectedSha256(await File.ReadAllTextAsync(shaPath, token));
                if (expected is null)
                {
                    throw new UpdateException("校验文件里没有 sha256");
                }

                EnsureSha256(exePath, expected);
            }

            return exePath;
        }
        catch (Exception ex)
        {
            try
            {
                Directory.Delete(dir, true);
            }
            catch (Exception cleanup)
            {
                AppLog.Exception("清理更新临时目录失败", cleanup);
            }

            if (ex is UpdateException)
            {
                throw;
            }

            AppLog.Exception("下载更新失败", ex);
            throw new UpdateException("下载失败: " + ex.Message, ex);
        }
    }

    public static string? ReadExpectedSha256(string text)
    {
        var match = Sha256Pattern.Match(text ?? "");
        return match.Success ? match.Value.ToLowerInvariant() : null;
    }

    public static void EnsureSha256(string path, string expected)
    {
        var actual = Sha256File(path);
        if (!actual.Equals(expected, StringComparison.OrdinalIgnoreCase))
        {
            throw new UpdateException("更新包校验失败：sha256 不一致");
        }
    }

    public static string Sha256File(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }

    public static string BuildApplyScript(PendingUpdate pending, int initialWaitSeconds = 2, int retryCount = 30)
    {
        var wait = Math.Max(0, initialWaitSeconds);
        var retries = Math.Max(1, retryCount);
        return "$ErrorActionPreference = 'Stop'\r\n"
            + $"$new = {PsQuote(pending.NewExe)}\r\n"
            + $"$target = {PsQuote(pending.TargetExe)}\r\n"
            + $"$bak = {PsQuote(pending.BackupExe)}\r\n"
            + $"$log = {PsQuote(pending.LogPath)}\r\n"
            + "function Write-Log([string]$message) {\r\n"
            + "  Add-Content -LiteralPath $log -Value (\"{0} {1}\" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $message)\r\n"
            + "}\r\n"
            + "Write-Log 'waiting target exit'\r\n"
            + $"Start-Sleep -Seconds {wait}\r\n"
            + $"for ($i = 0; $i -lt {retries}; $i++) {{\r\n"
            + "  try {\r\n"
            + "    if (Test-Path -LiteralPath $target) {\r\n"
            + "      Copy-Item -LiteralPath $target -Destination $bak -Force -ErrorAction SilentlyContinue\r\n"
            + "    }\r\n"
            + "    Copy-Item -LiteralPath $new -Destination $target -Force -ErrorAction Stop\r\n"
            + "    Write-Log 'update ok'\r\n"
            + "    exit 0\r\n"
            + "  } catch {\r\n"
            + "    Start-Sleep -Seconds 1\r\n"
            + "  }\r\n"
            + "}\r\n"
            + "Write-Log 'update failed: target locked'\r\n"
            + "exit 1\r\n";
    }

    public static PendingUpdate PrepareApply(string newExe, string targetExe, int initialWaitSeconds = 2, int retryCount = 30)
    {
        if (string.IsNullOrWhiteSpace(newExe) || !File.Exists(newExe))
        {
            throw new UpdateException("更新包不存在");
        }

        if (string.IsNullOrWhiteSpace(targetExe))
        {
            throw new UpdateException("当前不是打包后的 ExcelMergeFork.exe，无法原地更新。");
        }

        var pending = new PendingUpdate
        {
            ScriptPath = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(newExe))!, "apply_update.ps1"),
            NewExe = Path.GetFullPath(newExe),
            TargetExe = Path.GetFullPath(targetExe),
            BackupExe = Path.GetFullPath(targetExe) + ".bak",
            LogPath = Path.Combine(Path.GetDirectoryName(Path.GetFullPath(targetExe))!, "MergeExcelFork.update.log"),
        };
        var script = BuildApplyScript(pending, initialWaitSeconds, retryCount);
        File.WriteAllText(pending.ScriptPath, script, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
        return pending;
    }

    public static void LaunchApplyScript(PendingUpdate pending)
    {
        var powershell = Path.Combine(
            Environment.SystemDirectory,
            "WindowsPowerShell",
            "v1.0",
            "powershell.exe");
        var start = new ProcessStartInfo
        {
            FileName = powershell,
            UseShellExecute = false,
            CreateNoWindow = true,
            WorkingDirectory = Path.GetDirectoryName(pending.ScriptPath)!,
        };
        start.ArgumentList.Add("-NoLogo");
        start.ArgumentList.Add("-NoProfile");
        start.ArgumentList.Add("-ExecutionPolicy");
        start.ArgumentList.Add("Bypass");
        start.ArgumentList.Add("-WindowStyle");
        start.ArgumentList.Add("Hidden");
        start.ArgumentList.Add("-File");
        start.ArgumentList.Add(pending.ScriptPath);
        var process = Process.Start(start);
        if (process is null)
        {
            throw new UpdateException("无法启动更新脚本");
        }

        AppLog.Info("已启动更新脚本 " + pending.ScriptPath);
    }

    private static async Task DownloadFileAsync(
        string url,
        string path,
        string version,
        IProgress<UpdateProgress>? progress,
        CancellationToken token)
    {
        using var client = CreateClient(TimeSpan.FromMinutes(15));
        using var response = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead, token);
        if (!response.IsSuccessStatusCode)
        {
            throw new UpdateException($"下载失败: HTTP {(int)response.StatusCode}");
        }

        var total = response.Content.Headers.ContentLength ?? 0;
        await using var input = await response.Content.ReadAsStreamAsync(token);
        await using var output = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None, 128 * 1024, useAsync: true);
        var buffer = new byte[256 * 1024];
        long downloaded = 0;
        while (true)
        {
            var read = await input.ReadAsync(buffer, token);
            if (read == 0)
            {
                break;
            }

            await output.WriteAsync(buffer.AsMemory(0, read), token);
            downloaded += read;
            if (progress is not null && total > 0)
            {
                var percent = (int)Math.Min(100, downloaded * 100 / total);
                progress.Report(new UpdateProgress($"正在下载 v{version}... {percent}%"));
            }
            else
            {
                progress?.Report(new UpdateProgress($"正在下载 v{version}..."));
            }
        }
    }

    private static HttpClient CreateClient(TimeSpan timeout)
    {
        var client = new HttpClient { Timeout = timeout };
        client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", "ExcelMergeFork/" + AppVersion.Display);
        return client;
    }

    private static string PsQuote(string path) => "'" + path.Replace("'", "''") + "'";
}
