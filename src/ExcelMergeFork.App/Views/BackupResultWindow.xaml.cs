using System.Diagnostics;
using System.IO;
using System.Windows;
using ExcelMergeFork.Core.Backup;

namespace ExcelMergeFork.App.Views;

public partial class BackupResultWindow
{
    private readonly BackupInfo? _info;

    public BackupResultWindow(BackupInfo? info, string? error)
    {
        InitializeComponent();
        _info = info;
        if (error is not null)
        {
            Headline.Text = "备份失败";
            Detail.Text = error;
            return;
        }

        Headline.Text = "备份成功";
        Detail.Text = info?.Dir ?? "";
        FilesGrid.ItemsSource = info?.Files.Select(kv => new { Label = kv.Key, Path = kv.Value }).ToList();
    }

    private void OnOpen(object sender, RoutedEventArgs e)
    {
        if (_info is null || !Directory.Exists(_info.Dir))
        {
            return;
        }

        Process.Start(new ProcessStartInfo(_info.Dir) { UseShellExecute = true });
    }

    private void OnClose(object sender, RoutedEventArgs e) => Close();
}
