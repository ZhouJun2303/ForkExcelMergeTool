using System.Windows;
using System.Linq;
using ExcelMergeFork.App.Views;
using ExcelMergeFork.Core;
using ExcelMergeFork.Core.Backup;
using ExcelMergeFork.Core.Excel;
using ExcelMergeFork.Core.Git;
using ExcelMergeFork.Core.Routing;
using ExcelMergeFork.Core.Settings;
using Wpf.Ui.Appearance;

namespace ExcelMergeFork.App;

public partial class App : System.Windows.Application
{
    private void OnStartup(object sender, StartupEventArgs e)
    {
        var settings = AppSettingsStore.Load();
        ApplicationThemeManager.Apply(settings.DarkTheme ? ApplicationTheme.Dark : ApplicationTheme.Light);

        var request = LaunchArgs.Parse(e.Args);
        try
        {
            switch (request.Mode)
            {
                case LaunchMode.Settings:
                    new SettingsWindow().Show();
                    break;
                case LaunchMode.InstallFork:
                    ShowStatus("Fork 注入", ForkIntegration.Install(request.Files.FirstOrDefault()).Detail);
                    Shutdown(0);
                    break;
                case LaunchMode.UninstallFork:
                    ShowStatus("Fork 注入", ForkIntegration.Uninstall(request.Files.FirstOrDefault()).Detail);
                    Shutdown(0);
                    break;
                case LaunchMode.InstallGit:
                    ShowStatus("Git 注入", GitIntegration.Install(request.Files.FirstOrDefault()).Detail);
                    Shutdown(0);
                    break;
                case LaunchMode.UninstallGit:
                    ShowStatus("Git 注入", GitIntegration.Uninstall().Detail);
                    Shutdown(0);
                    break;
                case LaunchMode.GitDriver:
                    StartGitDriver(request);
                    break;
                case LaunchMode.Merge:
                    StartMerge(request);
                    break;
                case LaunchMode.Compare:
                    StartCompare(request);
                    break;
                default:
                    MessageBox.Show("参数无法识别。请从 Fork 打开，或直接双击进入设置中心。", "ExcelMergeFork");
                    new SettingsWindow().Show();
                    break;
            }
        }
        catch (Exception ex)
        {
            AppLog.Exception("启动失败", ex);
            MessageBox.Show(ex.Message, "ExcelMergeFork");
            Shutdown(1);
        }
    }

    private void StartMerge(LaunchRequest request)
    {
        var files = new[] { request.Local, request.Base, request.Remote, request.Merged };
        if (files.Any(string.IsNullOrWhiteSpace))
        {
            Shutdown(1);
            return;
        }

        var feature = ResolveFeature("merge", [
            ("本地 Local", request.Local!),
            ("基准 Base", request.Base!),
            ("线上 Remote", request.Remote!),
            ("输出 Merged", request.Merged!),
        ], files!);
        if (feature is null)
        {
            Shutdown(1);
            return;
        }

        if (feature == StartupFeature.BackupOnly)
        {
            ShowBackup(BackupService.CreateQuickBackup(
                [("local", request.Local!), ("base", request.Base!), ("remote", request.Remote!), ("merged", request.Merged!)],
                request.Merged!));
            return;
        }

        if (LaunchArgs.UnsupportedMergeDiff(files!).Count > 0)
        {
            MessageBox.Show(
                "合并对比模式当前只支持 " + ExcelFormats.MergeDiffExtensionText + "。宏文件请改用快速备份。",
                "ExcelMergeFork");
            Shutdown(1);
            return;
        }

        new MergeWindow(request.Local!, request.Base!, request.Remote!, request.Merged!).Show();
    }

    private void StartCompare(LaunchRequest request)
    {
        var files = new[] { request.Remote, request.Local };
        if (files.Any(string.IsNullOrWhiteSpace))
        {
            Shutdown(1);
            return;
        }

        var feature = ResolveFeature("compare", [
            ("A / 线上", request.Remote!),
            ("B / 本地", request.Local!),
        ], files!);
        if (feature is null)
        {
            Shutdown(1);
            return;
        }

        if (feature == StartupFeature.BackupOnly)
        {
            ShowBackup(BackupService.CreateQuickBackup([("a", request.Remote!), ("b", request.Local!)], request.Local!));
            return;
        }

        new DiffWindow(request.Local!, request.Remote!).Show();
    }

    private void StartGitDriver(LaunchRequest request)
    {
        if (request.Files.Count < 4)
        {
            Shutdown(1);
            return;
        }

        var prepared = GitMergeDriver.Prepare(request.Files[0], request.Files[1], request.Files[2], request.Files[3]);
        var context = GitMergeDriver.ContextPath(request.Files[1], request.Files[3]);
        var feature = ResolveFeature("git-driver", [
            ("BASE", prepared.BasePath),
            ("CURRENT", prepared.CurrentPath),
            ("OTHER", prepared.OtherPath),
            ("仓库", request.Files[3]),
        ], [prepared.BaseCopy, prepared.LocalCopy, prepared.RemoteCopy]);
        if (feature is null)
        {
            Shutdown(1);
            return;
        }

        if (feature == StartupFeature.BackupOnly)
        {
            ShowBackup(BackupService.CreateQuickBackup(
                [("base", prepared.BasePath), ("current", prepared.CurrentPath), ("other", prepared.OtherPath)],
                context));
            Shutdown(1);
            return;
        }

        // OnLastWindowClose calls Shutdown(0) before Window.Closed, which would
        // swallow cancel/X (exit 1). Git must wait for our explicit Shutdown.
        System.Windows.Application.Current.ShutdownMode = ShutdownMode.OnExplicitShutdown;
        var window = new MergeWindow(prepared.LocalCopy, prepared.BaseCopy, prepared.RemoteCopy, prepared.MergedCopy, prepared);
        window.Closed += (_, _) =>
            System.Windows.Application.Current.Shutdown(GitMergeDriver.WindowCloseExitCode(window.WriteBackSucceeded));
        window.Show();
    }

    private static string? ResolveFeature(string scene, IReadOnlyList<(string Role, string Path)> files, IEnumerable<string> checkPaths)
    {
        var feature = AppSettingsStore.Load().StartupFeatureValue;
        if (feature != StartupFeature.AskEachTime)
        {
            return feature;
        }

        var dialog = new StartupChoiceWindow(scene, files, LaunchArgs.UnsupportedMergeDiff(checkPaths));
        return dialog.ShowDialog() == true ? dialog.Choice : null;
    }

    private static void ShowBackup(BackupInfo info)
    {
        new BackupResultWindow(info, null).Show();
    }

    private static void ShowStatus(string title, string detail)
    {
        MessageBox.Show(detail, title);
    }
}
