using System.Diagnostics;
using ExcelMergeFork.Core;
using ExcelMergeFork.Core.Backup;
using ExcelMergeFork.Core.Excel;
using ExcelMergeFork.Core.Git;
using ExcelMergeFork.Core.Settings;
using ExcelMergeFork.Core.Update;

namespace ExcelMergeFork.Api;

public static class HostEndpoints
{
    public static void Map(RouteGroupBuilder api)
    {
        api.MapGet("/boot", () => Results.Ok(BootPayload(ExcelMergeHost.Desktop)));
        api.MapPost("/boot/choose", (ChooseBody? body, WorkSessionStore store) =>
            ApiRoutes.Guard(() =>
            {
                var desktop = ExcelMergeHost.Desktop ?? throw new InvalidOperationException("当前不是桌面启动");
                DesktopLaunch.Choose(desktop, store, body?.Choice);
                return Results.Ok(BootPayload(desktop));
            }));
        api.MapPost("/boot/cancel", () =>
        {
            var desktop = ExcelMergeHost.Desktop;
            if (desktop is not null)
            {
                desktop.ExitCode = 1;
            }

            return Results.Ok(new { close = desktop is not null, exitCode = 1 });
        });

        api.MapGet("/settings", () => Results.Ok(SettingsView(AppSettingsStore.Load())));
        api.MapPut("/settings", (SettingsBody? body) =>
            ApiRoutes.Guard(() => Results.Ok(SettingsView(ApplySettings(body)))));
        api.MapGet("/integrations", () => Results.Ok(IntegrationView()));
        api.MapPost("/integrations/fork/install", () => Results.Ok(IntegrationView(ForkIntegration.Install())));
        api.MapPost("/integrations/fork/uninstall", () => Results.Ok(IntegrationView(fork: ForkIntegration.Uninstall())));
        api.MapPost("/integrations/git/install", () => Results.Ok(IntegrationView(git: GitIntegration.Install())));
        api.MapPost("/integrations/git/uninstall", () => Results.Ok(IntegrationView(git: GitIntegration.Uninstall())));

        api.MapGet("/update/check", () => ApiRoutes.GuardAsync(CheckUpdateAsync));
        api.MapPost("/update/apply", () => ApiRoutes.GuardAsync(ApplyUpdateAsync));
        api.MapPost("/open", (OpenPathRequest? body) =>
            ApiRoutes.Guard(() =>
            {
                OpenPath(body?.Path);
                return Results.Ok(new { ok = true });
            }));

        api.MapPost("/merge/sessions/{id}/confirm", (string id, WorkSessionStore store) =>
            ApiRoutes.Guard(() => store.WithMerge(id, work =>
            {
                var result = work.Confirm();
                if (!result.Success)
                {
                    var message = result.Errors.Count == 0 ? result.Message : string.Join("；", result.Errors);
                    if (string.IsNullOrWhiteSpace(message))
                    {
                        message = "确认失败";
                    }

                    throw new InvalidOperationException(message);
                }

                return Results.Ok(new
                {
                    result.Success,
                    result.Message,
                    result.Staged,
                    close = ExcelMergeHost.Desktop is not null,
                });
            })));

        api.MapPost("/merge/sessions/{id}/ai/refresh", (string id, HttpRequest request, WorkSessionStore store) =>
            ApiRoutes.Guard(() => store.WithMerge(id, work =>
            {
                work.Ai.Refresh();
                return Results.Ok(work.Snapshot(PageQuery.From(request)));
            })));
        api.MapPost("/merge/sessions/{id}/ai/checks", (string id, AiCheckRequest? body, HttpRequest request, WorkSessionStore store) =>
            ApiRoutes.Guard(() => store.WithMerge(id, work =>
            {
                body ??= new AiCheckRequest();
                if (body.All is bool all)
                {
                    work.Ai.SetAll(all);
                }
                else
                {
                    work.Ai.SetChecked(body.Path, body.Checked);
                }

                return Results.Ok(work.Snapshot(PageQuery.From(request)));
            })));
        api.MapPost("/merge/sessions/{id}/ai/prompt", (string id, AiPromptRequest? body, WorkSessionStore store) =>
            ApiRoutes.Guard(() => store.WithMerge(id, work =>
            {
                var text = work.Ai.BuildPrompt(body?.Detail ?? false);
                return Results.Ok(new { text, status = work.Ai.StatusText });
            })));
    }

    public static void SaveMergeFlags(ExcelMergeFork.Core.Merge.MergeOptions options)
    {
        var settings = AppSettingsStore.Load();
        settings.SkipNewRows = options.SkipNewRows;
        settings.SkipNewColumns = options.SkipNewColumns;
        settings.DeleteMissingRows = options.DeleteMissingRows;
        settings.DeleteMissingColumns = options.DeleteMissingColumns;
        settings.AddNewSheets = options.AddNewSheets;
        settings.DeleteMissingSheets = options.DeleteMissingSheets;
        settings.ResolveConflicts = options.ResolveConflicts;
        AppSettingsStore.Save(settings);
    }

    private static object BootPayload(DesktopContext? desktop)
    {
        var settings = SettingsView(AppSettingsStore.Load());
        if (desktop is null)
        {
            return new
            {
                screen = "settings",
                host = "browser",
                mergeId = (string?)null,
                compareId = (string?)null,
                settings,
                version = AppVersion.Display,
                error = (string?)null,
                choice = new { title = "", notice = "", warn = "", canMerge = true, files = Array.Empty<BootFile>() },
                backup = new { ok = false, headline = "", detail = "", dir = (string?)null, files = Array.Empty<BootFile>() },
            };
        }

        return new
        {
            screen = desktop.Screen,
            host = "desktop",
            mergeId = desktop.MergeId,
            compareId = desktop.CompareId,
            settings,
            version = AppVersion.Display,
            error = desktop.Error,
            choice = new
            {
                title = desktop.ChoiceTitle,
                notice = desktop.ChoiceNotice,
                warn = desktop.ChoiceWarn,
                canMerge = desktop.CanMerge,
                files = desktop.ChoiceFiles,
            },
            backup = new
            {
                ok = desktop.BackupOk,
                headline = desktop.BackupHeadline,
                detail = desktop.BackupDetail,
                dir = desktop.BackupDir,
                files = desktop.BackupFiles,
            },
        };
    }

    private static object SettingsView(UserSettings settings) => new
    {
        settings.SkipNewRows,
        settings.SkipNewColumns,
        settings.DeleteMissingRows,
        settings.DeleteMissingColumns,
        settings.AddNewSheets,
        settings.DeleteMissingSheets,
        settings.ResolveConflicts,
        StartupFeature = settings.StartupFeatureValue,
        settings.BackupRootDir,
        settings.AutoOpenMerged,
        settings.AutoOpenCompare,
        settings.DarkTheme,
        settings.DiffFilter,
        Version = AppVersion.Display,
    };

    private static UserSettings ApplySettings(SettingsBody? body)
    {
        var settings = AppSettingsStore.Load();
        if (body is null)
        {
            return settings;
        }

        if (body.SkipNewRows is bool skipNewRows) settings.SkipNewRows = skipNewRows;
        if (body.SkipNewColumns is bool skipNewColumns) settings.SkipNewColumns = skipNewColumns;
        if (body.DeleteMissingRows is bool deleteMissingRows) settings.DeleteMissingRows = deleteMissingRows;
        if (body.DeleteMissingColumns is bool deleteMissingColumns) settings.DeleteMissingColumns = deleteMissingColumns;
        if (body.AddNewSheets is bool addNewSheets) settings.AddNewSheets = addNewSheets;
        if (body.DeleteMissingSheets is bool deleteMissingSheets) settings.DeleteMissingSheets = deleteMissingSheets;
        if (body.ResolveConflicts is bool resolveConflicts) settings.ResolveConflicts = resolveConflicts;
        if (body.StartupFeature is not null)
        {
            if (body.StartupFeature is not (StartupFeature.BackupOnly or StartupFeature.MergeDiff or StartupFeature.AskEachTime))
            {
                throw new InvalidOperationException("运行模式无效");
            }

            settings.StartupFeatureValue = body.StartupFeature;
        }

        if (body.BackupRootDir is not null) settings.BackupRootDir = body.BackupRootDir.Trim();
        if (body.AutoOpenMerged is bool autoOpenMerged) settings.AutoOpenMerged = autoOpenMerged;
        if (body.AutoOpenCompare is bool autoOpenCompare) settings.AutoOpenCompare = autoOpenCompare;
        if (body.DarkTheme is bool darkTheme) settings.DarkTheme = darkTheme;
        if (body.DiffFilter is not null)
        {
            foreach (var pair in body.DiffFilter)
            {
                settings.DiffFilter[pair.Key] = pair.Value;
            }
        }

        AppSettingsStore.Save(settings);
        return settings;
    }

    private static object IntegrationView(IntegrationStatus? fork = null, IntegrationStatus? git = null)
    {
        fork ??= ForkIntegration.Status();
        git ??= GitIntegration.Status();
        return new
        {
            fork = new { fork.Installed, fork.Detail, fork.Path },
            git = new { git.Installed, git.Detail, git.Path },
            toolPath = ForkIntegration.CurrentExecutable(),
            mergeArgs = "$LOCAL,$BASE,$REMOTE,$MERGED",
            diffArgs = "\"$REMOTE\" \"$LOCAL\"",
            driver = GitIntegration.DriverCommand(),
            driverArgs = "--git-merge-driver \"%O\" \"%A\" \"%B\" \"%P\"",
        };
    }

    private static async Task<IResult> CheckUpdateAsync()
    {
        UpdateInfo? info;
        try
        {
            info = await UpdateService.CheckAsync();
        }
        catch (UpdateException ex)
        {
            throw new InvalidOperationException(ex.Message);
        }

        if (info is null)
        {
            return Results.Ok(new
            {
                current = AppVersion.Display,
                found = false,
                isNewer = false,
                version = AppVersion.Display,
                message = "最新 Release 未找到 ExcelMergeFork.exe。",
            });
        }

        return Results.Ok(new
        {
            current = AppVersion.Display,
            found = true,
            info.IsNewer,
            info.Version,
            message = info.IsNewer
                ? "有新版本 v" + info.Version
                : "当前已是最新版本 v" + AppVersion.Display + "。",
        });
    }

    private static async Task<IResult> ApplyUpdateAsync()
    {
        UpdateInfo? info;
        try
        {
            info = await UpdateService.CheckAsync();
        }
        catch (UpdateException ex)
        {
            throw new InvalidOperationException(ex.Message);
        }

        if (info is null)
        {
            throw new InvalidOperationException("最新 Release 未找到 ExcelMergeFork.exe。");
        }

        if (!info.IsNewer)
        {
            return Results.Ok(new { close = false, message = "当前已是最新版本 v" + AppVersion.Display + "。" });
        }

        var target = UpdateService.CurrentExecutable();
        if (target is null)
        {
            throw new InvalidOperationException("当前不是打包后的 ExcelMergeFork.exe，无法原地更新。");
        }

        string newExe;
        try
        {
            newExe = await UpdateService.DownloadAsync(info, new Progress<UpdateProgress>(_ => { }));
            var pending = UpdateService.PrepareApply(newExe, target);
            UpdateService.LaunchApplyScript(pending);
        }
        catch (UpdateException ex)
        {
            throw new InvalidOperationException(ex.Message);
        }

        return Results.Ok(new
        {
            close = true,
            message = "更新包已下载。点击确定后会关闭当前窗口，并自动替换 ExcelMergeFork.exe。",
        });
    }

    private static void OpenPath(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            var root = BackupService.ResolveRoot(AppPaths.Home);
            Directory.CreateDirectory(root);
            Process.Start(new ProcessStartInfo(root) { UseShellExecute = true });
            return;
        }

        var path = RepoLayout.ResolvePath(raw);
        if (Directory.Exists(path))
        {
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            return;
        }

        if (File.Exists(path) && ExcelFormats.MergeDiffSupported(path))
        {
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
            return;
        }

        throw new InvalidOperationException("只能打开目录或 Excel 文件");
    }
}
