namespace ExcelMergeFork.Api;

public static class ApiRoutes
{
    public static void MapExcelApi(this WebApplication app)
    {
        var api = app.MapGroup("/api");
        api.MapGet("/health", () => Results.Ok(new { ok = true, name = "ExcelMergeFork" }));
        api.MapGet("/fixtures", () => Results.Ok(FixtureCatalog.Build()));
        HostEndpoints.Map(api);

        api.MapPost("/merge/sessions", (OpenMergeRequest? body, WorkSessionStore store, HttpRequest request) =>
            Guard(() =>
            {
                body ??= new OpenMergeRequest();
                var session = store.OpenMerge(
                    RepoLayout.RequireWorkbook(body.Local, "本地"),
                    RepoLayout.RequireWorkbook(body.Base, "基准"),
                    RepoLayout.RequireWorkbook(body.Remote, "线上"),
                    body.Merged,
                    body.Options,
                    body.Mode);
                return store.WithMerge(session.Id, work => Results.Ok(work.Snapshot(PageQuery.From(request))));
            }));

        api.MapPost("/merge/upload", async (HttpRequest request, WorkSessionStore store) =>
            await GuardAsync(async () =>
            {
                using var uploaded = await ReadUploadsAsync(request, ["local", "base", "remote"]);
                var session = store.OpenMerge(uploaded.Files["local"], uploaded.Files["base"], uploaded.Files["remote"], null, null, null);
                return store.WithMerge(session.Id, work => Results.Ok(work.Snapshot(PageQuery.From(request))));
            }));

        api.MapGet("/merge/sessions/{id}", (string id, HttpRequest request, WorkSessionStore store) =>
            Guard(() => store.WithMerge(id, work => Results.Ok(work.Snapshot(PageQuery.From(request))))));

        api.MapPut("/merge/sessions/{id}/options", (string id, MergeOptionsBody? body, HttpRequest request, WorkSessionStore store) =>
            Guard(() => store.WithMerge(id, work =>
            {
                body ??= new MergeOptionsBody();
                work.ApplyOptions(MergeWorkSession.ResolveOptions(work.Options, body, body.Mode));
                HostEndpoints.SaveMergeFlags(work.Options);
                return Results.Ok(work.Snapshot(PageQuery.From(request)));
            })));

        api.MapPost("/merge/sessions/{id}/choice", (string id, ChoiceRequest? body, HttpRequest request, WorkSessionStore store) =>
            Guard(() => store.WithMerge(id, work =>
            {
                body ??= new ChoiceRequest();
                work.Choose(body.Sheet, body.Key, body.Kind, body.Choice);
                return Results.Ok(work.Snapshot(PageQuery.From(request)));
            })));

        api.MapPost("/merge/sessions/{id}/sheets/select", (string id, SelectSheetRequest? body, HttpRequest request, WorkSessionStore store) =>
            Guard(() => store.WithMerge(id, work =>
            {
                work.SelectSheet(body?.Index ?? 0);
                return Results.Ok(work.Snapshot(PageQuery.From(request)));
            })));

        api.MapPost("/merge/sessions/{id}/sheets/complete", (string id, HttpRequest request, WorkSessionStore store) =>
            Guard(() => store.WithMerge(id, work =>
            {
                work.CompleteCurrent();
                return Results.Ok(work.Snapshot(PageQuery.From(request)));
            })));

        api.MapPost("/merge/sessions/{id}/sheets/previous", (string id, HttpRequest request, WorkSessionStore store) =>
            Guard(() => store.WithMerge(id, work =>
            {
                work.PreviousSheet();
                return Results.Ok(work.Snapshot(PageQuery.From(request)));
            })));

        api.MapPost("/merge/sessions/{id}/sheets/next", (string id, HttpRequest request, WorkSessionStore store) =>
            Guard(() => store.WithMerge(id, work =>
            {
                work.NextSheet();
                return Results.Ok(work.Snapshot(PageQuery.From(request)));
            })));

        api.MapPost("/merge/sessions/{id}/generate", (string id, GenerateRequest? body, HttpRequest request, WorkSessionStore store) =>
            Guard(() => store.WithMerge(id, work =>
            {
                work.Generate(body?.Merged, body?.CreateBackup ?? false);
                return Results.Ok(work.Snapshot(PageQuery.From(request)));
            })));

        api.MapGet("/merge/sessions/{id}/download", (string id, WorkSessionStore store) =>
            Guard(() => store.WithMerge(id, work =>
            {
                if (string.IsNullOrWhiteSpace(work.GeneratedPath) || !File.Exists(work.GeneratedPath))
                {
                    throw new InvalidOperationException("还没有合并结果");
                }

                return ExcelFile(work.GeneratedPath);
            })));

        api.MapDelete("/merge/sessions/{id}", (string id, WorkSessionStore store) =>
        {
            store.CloseMerge(id);
            return Results.Ok(new { ok = true });
        });

        api.MapPost("/compare/sessions", (OpenCompareRequest? body, HttpRequest request, WorkSessionStore store) =>
            Guard(() =>
            {
                body ??= new OpenCompareRequest();
                var session = store.OpenCompare(
                    RepoLayout.RequireWorkbook(body.Left, "线上"),
                    RepoLayout.RequireWorkbook(body.Right, "本地"),
                    body.IncludeSame);
                return store.WithCompare(session.Id, work => Results.Ok(work.Snapshot(SheetOf(request), SheetSpecified(request), QueryOf(request), StatusesOf(request), PageQuery.From(request))));
            }));

        api.MapPost("/compare/upload", async (HttpRequest request, WorkSessionStore store) =>
            await GuardAsync(async () =>
            {
                using var uploaded = await ReadUploadsAsync(request, ["left", "right"]);
                var includeSame = string.Equals(request.Query["includeSame"], "true", StringComparison.OrdinalIgnoreCase);
                var session = store.OpenCompare(uploaded.Files["left"], uploaded.Files["right"], includeSame);
                return store.WithCompare(session.Id, work => Results.Ok(work.Snapshot(null, false, null, null, PageQuery.From(request))));
            }));

        api.MapGet("/compare/sessions/{id}", (string id, HttpRequest request, WorkSessionStore store) =>
            Guard(() => store.WithCompare(id, work => Results.Ok(work.Snapshot(SheetOf(request), SheetSpecified(request), QueryOf(request), StatusesOf(request), PageQuery.From(request))))));

        api.MapPut("/compare/sessions/{id}/options", (string id, CompareOptionsBody? body, HttpRequest request, WorkSessionStore store) =>
            Guard(() => store.WithCompare(id, work =>
            {
                work.SetIncludeSame(body?.IncludeSame ?? false);
                return Results.Ok(work.Snapshot(SheetOf(request), SheetSpecified(request), QueryOf(request), StatusesOf(request), PageQuery.From(request)));
            })));

        api.MapPost("/compare/sessions/{id}/swap", (string id, HttpRequest request, WorkSessionStore store) =>
            Guard(() => store.WithCompare(id, work =>
            {
                work.Swap();
                return Results.Ok(work.Snapshot(SheetOf(request), SheetSpecified(request), QueryOf(request), StatusesOf(request), PageQuery.From(request)));
            })));

        api.MapGet("/compare/sessions/{id}/rows/{index:int}", (string id, int index, HttpRequest request, WorkSessionStore store) =>
            Guard(() => store.WithCompare(id, work =>
            {
                var changedOnly = !string.Equals(request.Query["changedOnly"], "false", StringComparison.OrdinalIgnoreCase);
                return Results.Ok(work.Detail(index, changedOnly));
            })));

        api.MapPost("/compare/sessions/{id}/export", (string id, WorkSessionStore store) =>
            Guard(() => store.WithCompare(id, work =>
            {
                var path = work.Export();
                return Results.Ok(new { path, downloadUrl = "/api/compare/sessions/" + id + "/download" });
            })));

        api.MapGet("/compare/sessions/{id}/download", (string id, WorkSessionStore store) =>
            Guard(() => store.WithCompare(id, work => ExcelFile(work.Export()))));

        api.MapDelete("/compare/sessions/{id}", (string id, WorkSessionStore store) =>
        {
            store.CloseCompare(id);
            return Results.Ok(new { ok = true });
        });
    }

    private static IResult ExcelFile(string path) =>
        Results.File(path, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", Path.GetFileName(path));

    private static bool SheetSpecified(HttpRequest request) => request.Query.ContainsKey("sheet");

    private static string? SheetOf(HttpRequest request)
    {
        var sheet = request.Query["sheet"].ToString();
        return string.IsNullOrWhiteSpace(sheet) ? null : sheet;
    }

    private static string? QueryOf(HttpRequest request)
    {
        var q = request.Query["q"].ToString();
        return string.IsNullOrWhiteSpace(q) ? null : q;
    }

    private static IReadOnlyCollection<string>? StatusesOf(HttpRequest request)
    {
        if (!request.Query.ContainsKey("status"))
        {
            return null;
        }

        return request.Query["status"].Where(value => !string.IsNullOrWhiteSpace(value)).Select(value => value!).ToArray();
    }

    private static async Task<UploadSet> ReadUploadsAsync(HttpRequest request, IReadOnlyList<string> fields)
    {
        if (!request.HasFormContentType)
        {
            throw new InvalidOperationException("请用表单上传工作簿");
        }

        var form = await request.ReadFormAsync();
        var dir = Path.Combine(RepoLayout.WorkRoot, "upload-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(dir);
        var saved = new Dictionary<string, string>(StringComparer.Ordinal);
        try
        {
            foreach (var field in fields)
            {
                var file = form.Files.GetFile(field);
                if (file is null || file.Length == 0)
                {
                    throw new InvalidOperationException("缺少上传文件：" + field);
                }

                var name = RepoLayout.SafeFileName(file.FileName);
                if (!name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) &&
                    !name.EndsWith(".xltx", StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(field + " 只支持 .xlsx / .xltx");
                }

                var path = Path.Combine(dir, field + "_" + name);
                await using var stream = File.Create(path);
                await file.CopyToAsync(stream);
                saved[field] = path;
            }

            return new UploadSet(dir, saved);
        }
        catch
        {
            try
            {
                Directory.Delete(dir, true);
            }
            catch
            {
                // 上传失败时尽量清掉半成品。
            }

            throw;
        }
    }

    private sealed class UploadSet : IDisposable
    {
        public UploadSet(string dir, Dictionary<string, string> files)
        {
            Dir = dir;
            Files = files;
        }

        public string Dir { get; }
        public Dictionary<string, string> Files { get; }

        public void Dispose()
        {
            try
            {
                if (Directory.Exists(Dir))
                {
                    Directory.Delete(Dir, true);
                }
            }
            catch
            {
                // 会话已经复制过工作簿，上传目录删不掉也不影响合并。
            }
        }
    }

    internal static IResult Guard(Func<IResult> action)
    {
        try
        {
            return action();
        }
        catch (Exception ex) when (ex is InvalidOperationException or FileNotFoundException or ArgumentException)
        {
            return Results.Json(new { error = ex.Message }, statusCode: StatusCodes.Status400BadRequest);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return Results.Json(new { error = ex.Message }, statusCode: StatusCodes.Status500InternalServerError);
        }
    }

    internal static async Task<IResult> GuardAsync(Func<Task<IResult>> action)
    {
        try
        {
            return await action();
        }
        catch (Exception ex) when (ex is InvalidOperationException or FileNotFoundException or ArgumentException)
        {
            return Results.Json(new { error = ex.Message }, statusCode: StatusCodes.Status400BadRequest);
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex);
            return Results.Json(new { error = ex.Message }, statusCode: StatusCodes.Status500InternalServerError);
        }
    }
}
