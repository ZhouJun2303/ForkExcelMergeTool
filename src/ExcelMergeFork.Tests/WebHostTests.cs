using System.Net.Http.Json;
using System.Text.Json;
using ClosedXML.Excel;
using ExcelMergeFork.Core.Excel;
using Microsoft.AspNetCore.Mvc.Testing;

namespace ExcelMergeFork.Tests;

public class WebHostTests : IClassFixture<WebApplicationFactory<Program>>
{
    private readonly HttpClient _client;

    public WebHostTests(WebApplicationFactory<Program> factory)
    {
        _client = factory.CreateClient();
    }

    [Fact]
    public async Task HomeAndFixtures_AreServed()
    {
        var html = await _client.GetStringAsync("/");
        Assert.Contains("ExcelMergeFork", html);
        var script = await _client.GetStringAsync("/app.js");
        Assert.Contains("Excel 三向合并", script);
        Assert.Contains("本地 对比 线上", script);
        Assert.Contains("确认并解决冲突", script);
        Assert.Contains("导出对比工作簿", script);

        var fixtures = await Ok(await _client.GetAsync("/api/fixtures"));
        Assert.Contains(fixtures.GetProperty("merges").EnumerateArray(), item => item.GetProperty("id").GetString() == "mode-d");
        Assert.Contains(fixtures.GetProperty("compares").EnumerateArray(), item => item.GetProperty("id").GetString() == "mode-d");
    }

    [Fact]
    public async Task MergeModeD_AutoRowsFlowThroughDownload()
    {
        var local = TestRepo.Fixture("mode_d_local.xlsx");
        var remote = TestRepo.Fixture("mode_d_remote.xlsx");
        var opened = await Ok(await _client.PostAsJsonAsync("/api/merge/sessions", new
        {
            local,
            @base = local,
            remote,
            mode = "D",
        }));
        var id = opened.GetProperty("id").GetString();
        Assert.False(string.IsNullOrWhiteSpace(id));
        try
        {
            opened = await SelectSheet(id!, opened, "Data");
            var rows = opened.GetProperty("items").EnumerateArray().ToList();
            Assert.Contains(rows, item => item.GetProperty("key").GetString() == "k1" && !item.GetProperty("canChoose").GetBoolean());
            Assert.Contains(rows, item => item.GetProperty("key").GetString() == "k2" && !item.GetProperty("canChoose").GetBoolean());

            for (var i = 0; i < 8 && !opened.GetProperty("wizard").GetProperty("canGenerate").GetBoolean(); i++)
            {
                opened = await Ok(await _client.PostAsync("/api/merge/sessions/" + id + "/sheets/complete", null));
            }

            Assert.True(opened.GetProperty("wizard").GetProperty("canGenerate").GetBoolean());
            opened = await Ok(await _client.PostAsJsonAsync("/api/merge/sessions/" + id + "/generate", new
            {
                createBackup = false,
            }));

            var path = opened.GetProperty("generatedPath").GetString();
            Assert.True(File.Exists(path));
            using (var book = new XLWorkbook(path))
            {
                var snap = SheetSnapshot.From(book.Worksheet("Data"));
                Assert.Equal(["k1", "remote1", "remote2"], snap.RowsByKey["k1"].Select(CellText.From).ToList());
                Assert.Equal(["k2", "remote_a", "remote_b"], snap.RowsByKey["k2"].Select(CellText.From).ToList());
            }

            var download = await _client.GetAsync(opened.GetProperty("downloadUrl").GetString());
            Assert.True(download.IsSuccessStatusCode, await download.Content.ReadAsStringAsync());
            Assert.Contains("spreadsheetml", download.Content.Headers.ContentType?.MediaType);
        }
        finally
        {
            await _client.DeleteAsync("/api/merge/sessions/" + id);
        }
    }

    [Fact]
    public async Task CompareModeD_ReturnsRowsWithoutWritingBesideSource()
    {
        var left = TestRepo.Fixture("mode_d_remote.xlsx");
        var right = TestRepo.Fixture("mode_d_local.xlsx");
        var beside = Path.Combine(TestRepo.TestData, "mode_d_remote_compare.xlsx");
        if (File.Exists(beside))
        {
            File.Delete(beside);
        }

        var opened = await Ok(await _client.PostAsJsonAsync("/api/compare/sessions", new
        {
            left,
            right,
        }));
        var id = opened.GetProperty("id").GetString();
        try
        {
            Assert.True(opened.GetProperty("total").GetInt32() > 0);
            Assert.Contains(opened.GetProperty("sheetNames").EnumerateArray(), sheet => sheet.GetString() == "Data");
            Assert.False(File.Exists(beside));

            var exported = await Ok(await _client.PostAsync("/api/compare/sessions/" + id + "/export", null));
            var exportPath = exported.GetProperty("path").GetString();
            Assert.True(File.Exists(exportPath));
            Assert.False(File.Exists(beside));
            Assert.StartsWith(Path.GetTempPath(), exportPath, StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            await _client.DeleteAsync("/api/compare/sessions/" + id);
        }
    }

    private async Task<JsonElement> SelectSheet(string id, JsonElement snapshot, string name)
    {
        var sheets = snapshot.GetProperty("wizard").GetProperty("sheets").EnumerateArray();
        var match = sheets.Single(sheet => sheet.GetProperty("name").GetString() == name);
        if (match.GetProperty("current").GetBoolean())
        {
            return snapshot;
        }

        return await Ok(await _client.PostAsJsonAsync("/api/merge/sessions/" + id + "/sheets/select", new
        {
            index = match.GetProperty("index").GetInt32(),
        }));
    }

    private static async Task<JsonElement> Ok(HttpResponseMessage response)
    {
        var text = await response.Content.ReadAsStringAsync();
        Assert.True(response.IsSuccessStatusCode, (int)response.StatusCode + " " + text);
        using var doc = JsonDocument.Parse(text);
        return doc.RootElement.Clone();
    }
}
