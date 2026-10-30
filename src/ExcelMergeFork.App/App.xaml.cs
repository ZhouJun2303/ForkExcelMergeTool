using System.Windows;
using ExcelMergeFork.Api;
using ExcelMergeFork.Core;
using ExcelMergeFork.Core.Git;
using ExcelMergeFork.Core.Routing;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;

namespace ExcelMergeFork.App;

public partial class App : System.Windows.Application
{
    private WebApplication? _host;
    private DesktopContext? _desktop;
    private LaunchRequest? _request;

    private async void OnStartup(object sender, StartupEventArgs e)
    {
        ShutdownMode = ShutdownMode.OnExplicitShutdown;
        var request = LaunchArgs.Parse(e.Args);
        _request = request;
        AppLog.Info("启动 mode=" + request.Mode + " files=" + string.Join(" | ", request.Files));
        try
        {
            if (HandleInstall(request))
            {
                return;
            }

            var desktop = new DesktopContext { Request = request };
            ExcelMergeHost.Desktop = desktop;
            _desktop = desktop;
            var host = ExcelMergeHost.Build([]);
            _host = host;
            await host.StartAsync();
            DesktopLaunch.Activate(desktop, host.Services.GetRequiredService<WorkSessionStore>());
            if (ShouldStopBeforeWindow(request, desktop))
            {
                MessageBox.Show(desktop.Error ?? "无法打开", "ExcelMergeFork");
                await host.StopAsync();
                Shutdown(desktop.ExitCode == 0 ? 1 : desktop.ExitCode);
                return;
            }

            var url = host.Urls.FirstOrDefault() ?? "http://127.0.0.1:5088";
            var shell = new ShellWindow(url, desktop.Screen);
            shell.Closed += async (_, _) =>
            {
                try
                {
                    await host.StopAsync();
                }
                catch (Exception ex)
                {
                    AppLog.Exception("停止页面服务失败", ex);
                }

                Shutdown(shell.RuntimeMissing ? 1 : ExitCodeOf());
            };
            shell.Show();
            shell.Activate();
        }
        catch (Exception ex)
        {
            AppLog.Exception("启动失败", ex);
            MessageBox.Show(ex.Message, "ExcelMergeFork");
            if (_host is not null)
            {
                try
                {
                    await _host.StopAsync();
                }
                catch
                {
                    // 启动失败时服务可能还没听端口。
                }
            }

            Shutdown(1);
        }
    }

    private static bool HandleInstall(LaunchRequest request)
    {
        switch (request.Mode)
        {
            case LaunchMode.InstallFork:
                MessageBox.Show(ForkIntegration.Install(request.Files.FirstOrDefault()).Detail, "Fork 注入");
                Current.Shutdown(0);
                return true;
            case LaunchMode.UninstallFork:
                MessageBox.Show(ForkIntegration.Uninstall(request.Files.FirstOrDefault()).Detail, "Fork 注入");
                Current.Shutdown(0);
                return true;
            case LaunchMode.InstallGit:
                MessageBox.Show(GitIntegration.Install(request.Files.FirstOrDefault()).Detail, "Git 注入");
                Current.Shutdown(0);
                return true;
            case LaunchMode.UninstallGit:
                MessageBox.Show(GitIntegration.Uninstall().Detail, "Git 注入");
                Current.Shutdown(0);
                return true;
            default:
                return false;
        }
    }

    private static bool ShouldStopBeforeWindow(LaunchRequest request, DesktopContext desktop) =>
        desktop.ExitCode != 0 &&
        desktop.Screen == "settings" &&
        desktop.MergeId is null &&
        desktop.CompareId is null &&
        request.Mode is LaunchMode.Merge or LaunchMode.Compare or LaunchMode.GitDriver;

    private int ExitCodeOf()
    {
        if (_request?.Mode == LaunchMode.GitDriver)
        {
            return GitMergeDriver.WindowCloseExitCode(_desktop?.WriteBackSucceeded ?? false);
        }

        return _desktop?.ExitCode ?? 0;
    }
}
