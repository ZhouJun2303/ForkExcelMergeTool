using System.Net.Http.Headers;
using System.Text.Json;

namespace ExcelMergeFork.Core.Update;

public sealed class UpdateInfo
{
    public required string Tag { get; init; }
    public required string Version { get; init; }
    public required string DownloadUrl { get; init; }
    public string? Sha256Url { get; init; }
    public bool IsNewer { get; init; }
}

public static class UpdateService
{
    public static async Task<UpdateInfo?> CheckAsync(CancellationToken token = default)
    {
        using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(12) };
        client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("ExcelMergeFork", AppVersion.Display));
        var url = $"https://api.github.com/repos/{AppConstants.GitHubRepo}/releases/latest";
        using var response = await client.GetAsync(url, token);
        response.EnsureSuccessStatusCode();
        using var doc = JsonDocument.Parse(await response.Content.ReadAsStringAsync(token));
        var root = doc.RootElement;
        var tag = root.GetProperty("tag_name").GetString() ?? "";
        var version = tag.TrimStart('v', 'V');
        string? download = null;
        string? sha = null;
        if (root.TryGetProperty("assets", out var assets))
        {
            foreach (var asset in assets.EnumerateArray())
            {
                var name = asset.GetProperty("name").GetString() ?? "";
                var href = asset.GetProperty("browser_download_url").GetString();
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

        if (string.IsNullOrWhiteSpace(download))
        {
            return null;
        }

        return new UpdateInfo
        {
            Tag = tag,
            Version = version,
            DownloadUrl = download,
            Sha256Url = sha,
            IsNewer = IsNewer(version, AppVersion.Display),
        };
    }

    public static bool IsNewer(string remote, string local)
    {
        static int[] Parts(string text) =>
            text.Split('.', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
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
}
