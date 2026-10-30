using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.Extensions.FileProviders;

namespace ExcelMergeFork.Api;

public static class ExcelMergeHost
{
    public static DesktopContext? Desktop { get; set; }

    public static WebApplication Build(string[] args)
    {
        var builder = WebApplication.CreateBuilder(args);
        if (Desktop is not null)
        {
            builder.WebHost.UseUrls("http://127.0.0.1:0");
        }

        builder.WebHost.ConfigureKestrel(options =>
        {
            options.Limits.MaxRequestBodySize = 256L * 1024 * 1024;
        });
        builder.Services.Configure<FormOptions>(options =>
        {
            options.MultipartBodyLengthLimit = 256L * 1024 * 1024;
        });
        builder.Services.ConfigureHttpJsonOptions(options =>
        {
            options.SerializerOptions.PropertyNamingPolicy = JsonNamingPolicy.CamelCase;
            options.SerializerOptions.PropertyNameCaseInsensitive = true;
            options.SerializerOptions.Converters.Add(new JsonStringEnumConverter());
        });
        builder.Services.AddSingleton<WorkSessionStore>();

        var app = builder.Build();
        RepoLayout.Locate([
            Path.GetDirectoryName(Environment.ProcessPath),
            app.Environment.ContentRootPath,
            AppContext.BaseDirectory,
            Directory.GetCurrentDirectory(),
        ]);
        Directory.CreateDirectory(RepoLayout.WorkRoot);

        app.Use(async (context, next) =>
        {
            if (HttpMethods.IsGet(context.Request.Method) ||
                HttpMethods.IsHead(context.Request.Method) ||
                HttpMethods.IsOptions(context.Request.Method))
            {
                await next();
                return;
            }

            var origin = context.Request.Headers.Origin.ToString();
            if (origin.Length > 0 && !RepoLayout.SameOrigin(origin, context.Request.Host))
            {
                context.Response.StatusCode = StatusCodes.Status403Forbidden;
                await context.Response.WriteAsJsonAsync(new { error = "只接受本机页面发起的修改" });
                return;
            }

            await next();
        });

        var webRoot = RepoLayout.WebRoot;
        if (Directory.Exists(webRoot))
        {
            var webFiles = new PhysicalFileProvider(webRoot);
            app.UseDefaultFiles(new DefaultFilesOptions { FileProvider = webFiles });
            app.UseStaticFiles(new StaticFileOptions
            {
                FileProvider = webFiles,
                OnPrepareResponse = ctx => ctx.Context.Response.Headers.CacheControl = "no-cache",
            });
        }

        var icons = Path.Combine(RepoLayout.RepoRoot, "Assets", "icons");
        if (Directory.Exists(icons))
        {
            app.UseStaticFiles(new StaticFileOptions
            {
                FileProvider = new PhysicalFileProvider(icons),
                RequestPath = "/assets",
                OnPrepareResponse = ctx => ctx.Context.Response.Headers.CacheControl = "no-cache",
            });
        }

        app.MapExcelApi();
        app.MapFallback(async context =>
        {
            if (context.Request.Path.StartsWithSegments("/api"))
            {
                context.Response.StatusCode = StatusCodes.Status404NotFound;
                await context.Response.WriteAsJsonAsync(new { error = "没有这个接口" });
                return;
            }

            var index = Path.Combine(webRoot, "index.html");
            if (!File.Exists(index))
            {
                context.Response.StatusCode = StatusCodes.Status404NotFound;
                await context.Response.WriteAsJsonAsync(new { error = "找不到 web/index.html" });
                return;
            }

            context.Response.ContentType = "text/html; charset=utf-8";
            context.Response.Headers.CacheControl = "no-cache";
            await context.Response.SendFileAsync(index);
        });

        app.Lifetime.ApplicationStarted.Register(() =>
        {
            var urls = app.Urls.Count == 0 ? "(动态端口)" : string.Join(", ", app.Urls);
            app.Logger.LogInformation("ExcelMergeFork：{Urls}", urls);
        });
        return app;
    }
}
