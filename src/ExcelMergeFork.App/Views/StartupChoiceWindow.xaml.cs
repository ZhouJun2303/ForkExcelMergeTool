using System.IO;
using System.Windows;
using ExcelMergeFork.Core;
using ExcelMergeFork.Core.Excel;
using ExcelMergeFork.Core.Settings;

namespace ExcelMergeFork.App.Views;

public partial class StartupChoiceWindow : Window
{
    public string? Choice { get; private set; }

    public StartupChoiceWindow(string scene, IReadOnlyList<(string Role, string Path)> files, IReadOnlyList<string> unsupported)
    {
        InitializeComponent();
        TitleText.Text = scene switch
        {
            "compare" => "Excel 二向对比",
            "git-driver" => "Git Excel 冲突",
            _ => "Excel 三向合并",
        };
        NoticeText.Text = scene switch
        {
            "compare" => "大表建议先快速备份，小表可进入对比窗口。",
            "git-driver" => "快速备份会保留冲突未解决；合并确认后 Git 才会继续。",
            _ => "大表建议先快速备份，小表可进入合并窗口。",
        };
        FilesGrid.ItemsSource = files.Select(f => new
        {
            f.Role,
            Name = System.IO.Path.GetFileName(f.Path),
            Size = FormatSize(f.Path),
            f.Path,
        }).ToList();
        if (unsupported.Count > 0)
        {
            WarnText.Text = "这些后缀暂不支持合并/对比：" + string.Join("、", unsupported.Select(System.IO.Path.GetFileName))
                            + "。支持 " + ExcelFormats.MergeDiffExtensionText + "。";
            MergeButton.IsEnabled = false;
        }
    }

    private void OnBackup(object sender, RoutedEventArgs e) => Accept(StartupFeature.BackupOnly);

    private void OnMerge(object sender, RoutedEventArgs e) => Accept(StartupFeature.MergeDiff);

    private void OnCancel(object sender, RoutedEventArgs e)
    {
        AppLog.Info("启动选择：取消");
        Choice = null;
        TrySetDialogResult(false);
        Close();
    }

    public void Accept(string choice)
    {
        Choice = choice;
        AppLog.Info("启动选择：" + choice);
        TrySetDialogResult(true);
        if (IsVisible)
        {
            Close();
        }
    }

    private void TrySetDialogResult(bool value)
    {
        try
        {
            DialogResult = value;
        }
        catch (InvalidOperationException)
        {
            // Shown with Show() instead of ShowDialog(); Close() is enough.
        }
    }

    private static string FormatSize(string path)
    {
        try
        {
            if (!File.Exists(path))
            {
                return "不存在";
            }

            var size = (double)new FileInfo(path).Length;
            string[] units = ["B", "KB", "MB", "GB"];
            var i = 0;
            while (size >= 1024 && i < units.Length - 1)
            {
                size /= 1024;
                i++;
            }

            return i == 0 ? $"{(int)size} {units[i]}" : $"{size:0.0} {units[i]}";
        }
        catch
        {
            return "无法读取";
        }
    }
}
