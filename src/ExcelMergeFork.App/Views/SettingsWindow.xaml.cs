using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using ExcelMergeFork.Core;
using ExcelMergeFork.Core.Backup;
using ExcelMergeFork.Core.Git;
using ExcelMergeFork.Core.Settings;
using ExcelMergeFork.Core.Update;
using Microsoft.Win32;
using Wpf.Ui.Appearance;

namespace ExcelMergeFork.App.Views;

public partial class SettingsWindow
{
    private UserSettings _settings = AppSettingsStore.Load();
    private bool _updateBusy;

    public SettingsWindow()
    {
        InitializeComponent();
        Title = $"ExcelMergeFork 设置中心 v{AppVersion.Display}";
        // ListBoxItem.IsSelected in XAML fires SelectionChanged during InitializeComponent,
        // before PageHost exists. Load the first page only after the tree is ready.
        ShowPage("general");
    }

    private void OnNavChanged(object sender, SelectionChangedEventArgs e)
    {
        if (Nav.SelectedItem is ListBoxItem item && item.Tag is string tag)
        {
            ShowPage(tag);
        }
    }

    private void ShowPage(string tag)
    {
        if (PageHost is null)
        {
            return;
        }

        _settings = AppSettingsStore.Load();
        PageHost.Children.Clear();
        switch (tag)
        {
            case "backup":
                BuildBackup();
                break;
            case "fork":
                BuildFork();
                break;
            case "git":
                BuildGit();
                break;
            case "update":
                BuildUpdate();
                break;
            default:
                BuildGeneral();
                break;
        }
    }

    private void BuildGeneral()
    {
        AddTitle("默认运行模式", "从 Fork、命令行或 Git 传入文件时的默认行为。");
        AddRadio("快速备份模式", StartupFeature.BackupOnly, "只备份输入文件，不生成合并或对比结果。");
        AddRadio("合并对比模式", StartupFeature.MergeDiff, "打开完整三向合并或二向对比窗口。");
        AddRadio("每次询问", StartupFeature.AskEachTime, "每次先选择备份还是合并/对比。适合大表。");
        AddSeparator();
        var dark = new CheckBox { Content = "深色主题", IsChecked = _settings.DarkTheme, Margin = new Thickness(0, 8, 0, 0) };
        dark.Checked += (_, _) => SetTheme(true);
        dark.Unchecked += (_, _) => SetTheme(false);
        PageHost.Children.Add(dark);
    }

    private void BuildBackup()
    {
        AddTitle("备份根目录", "留空则使用目标文件同目录下的 MergeExcelBackup。");
        var box = new TextBox { Text = _settings.BackupRootDir, Margin = new Thickness(0, 12, 0, 8) };
        var row = new StackPanel { Orientation = Orientation.Horizontal };
        var pick = new Button { Content = "选择目录", Margin = new Thickness(0, 0, 8, 0) };
        pick.Click += (_, _) =>
        {
            var dialog = new OpenFolderDialog { Title = "选择备份根目录" };
            if (dialog.ShowDialog() == true)
            {
                box.Text = dialog.FolderName;
                _settings.BackupRootDir = dialog.FolderName;
                AppSettingsStore.Save(_settings);
            }
        };
        var save = new Button { Content = "保存" };
        save.Click += (_, _) =>
        {
            _settings.BackupRootDir = box.Text.Trim();
            AppSettingsStore.Save(_settings);
            System.Windows.MessageBox.Show("备份根目录已保存。");
        };
        var open = new Button { Content = "打开目录", Margin = new Thickness(8, 0, 0, 0) };
        open.Click += (_, _) =>
        {
            var root = string.IsNullOrWhiteSpace(box.Text) ? BackupService.ResolveRoot(AppPaths.Home) : box.Text;
            Directory.CreateDirectory(root);
            Process.Start(new ProcessStartInfo(root) { UseShellExecute = true });
        };
        row.Children.Add(pick);
        row.Children.Add(save);
        row.Children.Add(open);
        PageHost.Children.Add(box);
        PageHost.Children.Add(row);
    }

    private void BuildFork()
    {
        var status = ForkIntegration.Status();
        AddTitle("Fork 一键注入", "安装后，Fork 的 Merge Tool 和 External Diff Tool 会指向本工具。请先关闭 Fork。");
        PageHost.Children.Add(new TextBlock { Text = status.Detail, Margin = new Thickness(0, 12, 0, 12), TextWrapping = TextWrapping.Wrap });
        var row = new StackPanel { Orientation = Orientation.Horizontal };
        var install = new Button { Content = "安装注入", Margin = new Thickness(0, 0, 8, 0) };
        install.Click += (_, _) => { ForkIntegration.Install(); ShowPage("fork"); };
        var uninstall = new Button { Content = "移除注入" };
        uninstall.Click += (_, _) => { ForkIntegration.Uninstall(); ShowPage("fork"); };
        row.Children.Add(install);
        row.Children.Add(uninstall);
        PageHost.Children.Add(row);
        AddCopy("工具路径", ForkIntegration.CurrentExecutable());
        AddCopy("合并参数", "$LOCAL,$BASE,$REMOTE,$MERGED");
        AddCopy("对比参数", "\"$REMOTE\" \"$LOCAL\"");
    }

    private void BuildGit()
    {
        var status = GitIntegration.Status();
        AddTitle("全局 Git 注入", "安装后，任意 Git 工具遇到常见 Excel 后缀冲突都会调用本工具。");
        PageHost.Children.Add(new TextBlock { Text = status.Detail, Margin = new Thickness(0, 12, 0, 12), TextWrapping = TextWrapping.Wrap });
        var row = new StackPanel { Orientation = Orientation.Horizontal };
        var install = new Button { Content = "安装注入", Margin = new Thickness(0, 0, 8, 0) };
        install.Click += (_, _) => { GitIntegration.Install(); ShowPage("git"); };
        var uninstall = new Button { Content = "移除注入" };
        uninstall.Click += (_, _) => { GitIntegration.Uninstall(); ShowPage("git"); };
        row.Children.Add(install);
        row.Children.Add(uninstall);
        PageHost.Children.Add(row);
        AddCopy("driver 值", GitIntegration.DriverCommand());
        AddCopy("driver 参数", "--git-merge-driver \"%O\" \"%A\" \"%B\" \"%P\"");
    }

    private void BuildUpdate()
    {
        AddTitle("程序更新", "检查 GitHub Releases。不会自动替换，必须确认后才更新。");
        var state = new TextBlock
        {
            Text = $"当前版本 v{AppVersion.Display}",
            Margin = new Thickness(0, 12, 0, 12),
            TextWrapping = TextWrapping.Wrap,
        };
        var check = new Button
        {
            Content = "检查更新",
            HorizontalAlignment = HorizontalAlignment.Left,
            Padding = new Thickness(16, 6, 16, 6),
        };
        check.Click += async (_, _) => await CheckForUpdateAsync(check, state);
        PageHost.Children.Add(state);
        PageHost.Children.Add(check);
    }

    private async Task CheckForUpdateAsync(Button check, TextBlock state)
    {
        if (_updateBusy)
        {
            return;
        }

        _updateBusy = true;
        check.IsEnabled = false;
        check.Content = "检查中...";
        state.Text = "正在检查...";
        AppLog.Info("点击检查更新");
        try
        {
            UpdateInfo? info;
            try
            {
                info = await UpdateService.CheckAsync();
            }
            catch (Exception ex)
            {
                state.Text = "检查失败";
                AppLog.Exception("检查更新失败", ex);
                ShowUpdateDialog(ex.Message, "检查更新失败", MessageBoxImage.Error);
                return;
            }

            if (info is null)
            {
                state.Text = "未找到发布包";
                ShowUpdateDialog("最新 Release 未找到 ExcelMergeFork.exe。", "检查更新", MessageBoxImage.Warning);
                return;
            }

            if (!info.IsNewer)
            {
                state.Text = $"已是最新 v{info.Version}";
                ShowUpdateDialog($"当前已是最新版本 v{AppVersion.Display}。", "检查更新", MessageBoxImage.Information);
                return;
            }

            state.Text = $"有新版本 v{info.Version}";
            check.Content = $"有新版本 v{info.Version}";
            var confirm = ShowUpdateDialog(
                $"发现新版本 v{info.Version}，当前版本 v{AppVersion.Display}。\n\n"
                + "点击确定后会下载新版 exe；下载完成后需要关闭当前窗口，工具会在退出后替换文件。\n"
                + "如果正在合并冲突，建议先完成或取消当前合并后再更新。",
                "发现新版本",
                MessageBoxImage.Question,
                MessageBoxButton.OKCancel);
            if (confirm != MessageBoxResult.OK)
            {
                return;
            }

            var target = UpdateService.CurrentExecutable();
            if (target is null)
            {
                ShowUpdateDialog("当前不是打包后的 ExcelMergeFork.exe，无法原地更新。", "更新提示", MessageBoxImage.Information);
                return;
            }

            check.Content = "下载中...";
            var progress = new Progress<UpdateProgress>(p => state.Text = p.Status);
            string newExe;
            try
            {
                newExe = await UpdateService.DownloadAsync(info, progress);
            }
            catch (Exception ex)
            {
                state.Text = "下载失败";
                AppLog.Exception("下载更新失败", ex);
                ShowUpdateDialog(ex.Message, "更新失败", MessageBoxImage.Error);
                return;
            }

            PendingUpdate pending;
            try
            {
                pending = UpdateService.PrepareApply(newExe, target);
            }
            catch (Exception ex)
            {
                state.Text = "更新失败";
                AppLog.Exception("准备更新脚本失败", ex);
                ShowUpdateDialog(ex.Message, "更新失败", MessageBoxImage.Error);
                return;
            }

            state.Text = "下载完成，等待替换...";
            ShowUpdateDialog(
                "更新包已下载。点击确定后会关闭当前窗口，并自动替换 ExcelMergeFork.exe。",
                "更新已准备好",
                MessageBoxImage.Information);
            try
            {
                UpdateService.LaunchApplyScript(pending);
            }
            catch (Exception ex)
            {
                state.Text = "更新失败";
                AppLog.Exception("启动更新脚本失败", ex);
                ShowUpdateDialog(ex.Message, "更新失败", MessageBoxImage.Error);
                return;
            }

            System.Windows.Application.Current.Shutdown();
        }
        finally
        {
            _updateBusy = false;
            if (check.Content as string == "检查中..." || check.Content as string == "下载中...")
            {
                check.Content = "检查更新";
            }

            check.IsEnabled = true;
        }
    }

    private MessageBoxResult ShowUpdateDialog(
        string message,
        string title,
        MessageBoxImage image,
        MessageBoxButton buttons = MessageBoxButton.OK)
    {
        return System.Windows.MessageBox.Show(this, message, title, buttons, image);
    }

    private void AddTitle(string title, string subtitle)
    {
        PageHost.Children.Add(new TextBlock { Text = title, FontSize = 22, FontWeight = FontWeights.SemiBold });
        PageHost.Children.Add(new TextBlock
        {
            Text = subtitle,
            Opacity = 0.7,
            TextWrapping = TextWrapping.Wrap,
            Margin = new Thickness(0, 6, 0, 8),
        });
    }

    private void AddRadio(string title, string value, string hint)
    {
        var radio = new RadioButton
        {
            Content = title,
            GroupName = "startup",
            IsChecked = _settings.StartupFeatureValue == value,
            Margin = new Thickness(0, 10, 0, 0),
        };
        radio.Checked += (_, _) =>
        {
            _settings.StartupFeatureValue = value;
            AppSettingsStore.Save(_settings);
        };
        PageHost.Children.Add(radio);
        PageHost.Children.Add(new TextBlock { Text = hint, Opacity = 0.65, Margin = new Thickness(24, 2, 0, 0), TextWrapping = TextWrapping.Wrap });
    }

    private void AddSeparator() => PageHost.Children.Add(new Separator { Margin = new Thickness(0, 20, 0, 12) });

    private void AddCopy(string label, string value)
    {
        PageHost.Children.Add(new TextBlock { Text = label, Margin = new Thickness(0, 16, 0, 4), FontWeight = FontWeights.SemiBold });
        var row = new DockPanel();
        var copy = new Button { Content = "复制", Margin = new Thickness(8, 0, 0, 0) };
        DockPanel.SetDock(copy, Dock.Right);
        copy.Click += (_, _) => Clipboard.SetText(value);
        row.Children.Add(copy);
        row.Children.Add(new TextBox { Text = value, IsReadOnly = true });
        PageHost.Children.Add(row);
    }

    private void SetTheme(bool dark)
    {
        _settings.DarkTheme = dark;
        AppSettingsStore.Save(_settings);
        ApplicationThemeManager.Apply(dark ? ApplicationTheme.Dark : ApplicationTheme.Light);
    }
}
