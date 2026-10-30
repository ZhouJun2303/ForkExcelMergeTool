using ExcelMergeFork.Core;
using ExcelMergeFork.Core.Backup;
using ExcelMergeFork.Core.Excel;
using ExcelMergeFork.Core.Git;
using ExcelMergeFork.Core.Routing;
using ExcelMergeFork.Core.Settings;

namespace ExcelMergeFork.Api;

public static class DesktopLaunch
{
    public static void Activate(DesktopContext desktop, WorkSessionStore store)
    {
        var request = desktop.Request;
        switch (request.Mode)
        {
            case LaunchMode.Merge:
                ActivateMerge(desktop, store, request);
                break;
            case LaunchMode.Compare:
                ActivateCompare(desktop, store, request);
                break;
            case LaunchMode.GitDriver:
                ActivateGitDriver(desktop, store, request);
                break;
            default:
                desktop.Screen = "settings";
                if (request.Mode == LaunchMode.Invalid)
                {
                    desktop.Error = "参数无法识别。请从 Fork 打开，或直接进入设置中心。";
                }
                break;
        }
    }

    public static void Choose(DesktopContext desktop, WorkSessionStore store, string? choice)
    {
        if (choice == StartupFeature.BackupOnly)
        {
            CreatePendingBackup(desktop);
            desktop.Screen = "backup";
            return;
        }

        if (choice != StartupFeature.MergeDiff)
        {
            throw new InvalidOperationException("请选择快速备份或进入合并/对比");
        }

        if (!desktop.CanMerge)
        {
            throw new InvalidOperationException(desktop.ChoiceWarn);
        }

        OpenResolved(desktop, store);
    }

    private static void ActivateMerge(DesktopContext desktop, WorkSessionStore store, LaunchRequest request)
    {
        var files = new[] { request.Local, request.Base, request.Remote, request.Merged };
        if (files.Any(string.IsNullOrWhiteSpace))
        {
            desktop.ExitCode = 1;
            desktop.Error = "合并参数不完整";
            desktop.Screen = "settings";
            return;
        }

        var listed = new (string Role, string Path)[]
        {
            ("本地 Local", request.Local!),
            ("基准 Base", request.Base!),
            ("线上 Remote", request.Remote!),
            ("输出 Merged", request.Merged!),
        };
        if (!ResolveFeature(desktop, "merge", listed, files!))
        {
            return;
        }

        if (desktop.Screen == "choose")
        {
            return;
        }

        if (AppSettingsStore.Load().StartupFeatureValue == StartupFeature.BackupOnly)
        {
            desktop.BackupOk = true;
            FillBackup(desktop, BackupService.CreateQuickBackup(
                [("local", request.Local!), ("base", request.Base!), ("remote", request.Remote!), ("merged", request.Merged!)],
                request.Merged!));
            desktop.Screen = "backup";
            return;
        }

        var unsupported = LaunchArgs.UnsupportedMergeDiff(files!);
        if (unsupported.Count > 0)
        {
            desktop.ExitCode = 1;
            desktop.Error = "合并对比模式当前只支持 " + ExcelFormats.MergeDiffExtensionText + "。宏文件请改用快速备份。";
            desktop.Screen = "settings";
            return;
        }

        var session = store.OpenMerge(request.Local!, request.Base!, request.Remote!, request.Merged, null, null);
        desktop.MergeId = session.Id;
        desktop.Screen = "merge";
    }

    private static void ActivateCompare(DesktopContext desktop, WorkSessionStore store, LaunchRequest request)
    {
        var files = new[] { request.Remote, request.Local };
        if (files.Any(string.IsNullOrWhiteSpace))
        {
            desktop.ExitCode = 1;
            desktop.Error = "对比参数不完整";
            desktop.Screen = "settings";
            return;
        }

        var listed = new (string Role, string Path)[]
        {
            ("A / 线上", request.Remote!),
            ("B / 本地", request.Local!),
        };
        if (!ResolveFeature(desktop, "compare", listed, files!))
        {
            return;
        }

        if (desktop.Screen == "choose")
        {
            return;
        }

        if (AppSettingsStore.Load().StartupFeatureValue == StartupFeature.BackupOnly)
        {
            FillBackup(desktop, BackupService.CreateQuickBackup([("a", request.Remote!), ("b", request.Local!)], request.Local!));
            desktop.Screen = "backup";
            return;
        }

        var session = store.OpenCompare(request.Remote!, request.Local!, false);
        desktop.CompareId = session.Id;
        desktop.Screen = "compare";
    }

    private static void ActivateGitDriver(DesktopContext desktop, WorkSessionStore store, LaunchRequest request)
    {
        if (request.Files.Count < 4)
        {
            desktop.ExitCode = 1;
            desktop.Error = "Git driver 参数不完整";
            desktop.Screen = "settings";
            return;
        }

        var prepared = GitMergeDriver.Prepare(request.Files[0], request.Files[1], request.Files[2], request.Files[3]);
        desktop.Driver = prepared;
        var listed = new (string Role, string Path)[]
        {
            ("BASE", prepared.BasePath),
            ("CURRENT", prepared.CurrentPath),
            ("OTHER", prepared.OtherPath),
            ("仓库", request.Files[3]),
        };
        if (!ResolveFeature(desktop, "git-driver", listed, [prepared.BaseCopy, prepared.LocalCopy, prepared.RemoteCopy]))
        {
            return;
        }

        if (desktop.Screen == "choose")
        {
            return;
        }

        if (AppSettingsStore.Load().StartupFeatureValue == StartupFeature.BackupOnly)
        {
            var context = GitMergeDriver.ContextPath(request.Files[1], request.Files[3]);
            FillBackup(desktop, BackupService.CreateQuickBackup(
                [("base", prepared.BasePath), ("current", prepared.CurrentPath), ("other", prepared.OtherPath)],
                context));
            desktop.Screen = "backup";
            desktop.ExitCode = GitMergeDriver.WindowCloseExitCode(false);
            return;
        }

        OpenGitMerge(desktop, store, prepared);
    }

    private static void OpenResolved(DesktopContext desktop, WorkSessionStore store)
    {
        var request = desktop.Request;
        if (request.Mode == LaunchMode.Compare)
        {
            var session = store.OpenCompare(request.Remote!, request.Local!, false);
            desktop.CompareId = session.Id;
            desktop.Screen = "compare";
            return;
        }

        if (request.Mode == LaunchMode.GitDriver)
        {
            if (desktop.Driver is null)
            {
                throw new InvalidOperationException("Git driver 尚未准备");
            }

            OpenGitMerge(desktop, store, desktop.Driver);
            return;
        }

        var sessionMerge = store.OpenMerge(request.Local!, request.Base!, request.Remote!, request.Merged, null, null);
        desktop.MergeId = sessionMerge.Id;
        desktop.Screen = "merge";
    }

    private static void OpenGitMerge(DesktopContext desktop, WorkSessionStore store, GitDriverRequest prepared)
    {
        var session = store.OpenMerge(prepared.LocalCopy, prepared.BaseCopy, prepared.RemoteCopy, prepared.MergedCopy, null, null);
        store.WithMerge(session.Id, work =>
        {
            work.AttachDriver(prepared);
            return true;
        });
        desktop.MergeId = session.Id;
        desktop.Screen = "merge";
    }

    private static void CreatePendingBackup(DesktopContext desktop)
    {
        var request = desktop.Request;
        if (request.Mode == LaunchMode.Compare)
        {
            FillBackup(desktop, BackupService.CreateQuickBackup([("a", request.Remote!), ("b", request.Local!)], request.Local!));
            return;
        }

        if (request.Mode == LaunchMode.GitDriver && desktop.Driver is not null)
        {
            var prepared = desktop.Driver;
            var context = GitMergeDriver.ContextPath(request.Files[1], request.Files[3]);
            FillBackup(desktop, BackupService.CreateQuickBackup(
                [("base", prepared.BasePath), ("current", prepared.CurrentPath), ("other", prepared.OtherPath)],
                context));
            desktop.ExitCode = GitMergeDriver.WindowCloseExitCode(false);
            return;
        }

        FillBackup(desktop, BackupService.CreateQuickBackup(
            [("local", request.Local!), ("base", request.Base!), ("remote", request.Remote!), ("merged", request.Merged!)],
            request.Merged!));
    }

    private static bool ResolveFeature(
        DesktopContext desktop,
        string scene,
        IReadOnlyList<(string Role, string Path)> files,
        IEnumerable<string> checkPaths)
    {
        var feature = AppSettingsStore.Load().StartupFeatureValue;
        if (feature == StartupFeature.AskEachTime)
        {
            var unsupported = LaunchArgs.UnsupportedMergeDiff(checkPaths);
            desktop.Screen = "choose";
            desktop.ChoiceTitle = scene switch
            {
                "compare" => "Excel 二向对比",
                "git-driver" => "Git Excel 冲突",
                _ => "Excel 三向合并",
            };
            desktop.ChoiceNotice = scene switch
            {
                "compare" => "大表建议先快速备份，小表可进入对比窗口。",
                "git-driver" => "快速备份会保留冲突未解决；合并确认后 Git 才会继续。",
                _ => "大表建议先快速备份，小表可进入合并窗口。",
            };
            desktop.ChoiceFiles = files.Select(file => new BootFile
            {
                Role = file.Role,
                Name = Path.GetFileName(file.Path),
                Size = FormatSize(file.Path),
                Path = file.Path,
            }).ToList();
            desktop.CanMerge = unsupported.Count == 0;
            desktop.ChoiceWarn = unsupported.Count == 0
                ? ""
                : "这些后缀暂不支持合并/对比：" + string.Join("、", unsupported.Select(Path.GetFileName))
                  + "。支持 " + ExcelFormats.MergeDiffExtensionText + "。";
            return true;
        }

        return true;
    }

    private static void FillBackup(DesktopContext desktop, BackupInfo info)
    {
        desktop.BackupOk = true;
        desktop.BackupHeadline = "备份成功";
        desktop.BackupDetail = info.Dir;
        desktop.BackupDir = info.Dir;
        desktop.BackupFiles = info.Files.Select(pair => new BootFile
        {
            Role = pair.Key,
            Name = Path.GetFileName(pair.Value),
            Size = "",
            Path = pair.Value,
        }).ToList();
    }

    private static string FormatSize(string path)
    {
        if (!File.Exists(path))
        {
            return "";
        }

        var length = new FileInfo(path).Length;
        if (length < 1024)
        {
            return length + " B";
        }

        if (length < 1024 * 1024)
        {
            return (length / 1024d).ToString("0.#") + " KB";
        }

        return (length / 1024d / 1024d).ToString("0.#") + " MB";
    }
}
