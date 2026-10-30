using System.IO;
using System.Text.Json;
using System.Windows;
using ExcelMergeFork.Core;
using Microsoft.Web.WebView2.Core;
using Microsoft.Win32;

namespace ExcelMergeFork.App;

public partial class ShellWindow : Window
{
    private readonly string _url;
    private bool _ready;

    public ShellWindow(string url, string screen)
    {
        InitializeComponent();
        _url = url;
        ApplySize(screen);
        Loaded += OnLoaded;
    }

    public bool RuntimeMissing { get; private set; }

    private async void OnLoaded(object sender, RoutedEventArgs e)
    {
        if (_ready)
        {
            return;
        }

        _ready = true;
        try
        {
            var folder = Path.Combine(Path.GetTempPath(), "ExcelMergeFork", "webview");
            Directory.CreateDirectory(folder);
            var env = await CoreWebView2Environment.CreateAsync(null, folder);
            await View.EnsureCoreWebView2Async(env);
            View.CoreWebView2.Settings.IsStatusBarEnabled = false;
            View.CoreWebView2.PermissionRequested += (_, args) =>
            {
                if (args.PermissionKind == CoreWebView2PermissionKind.ClipboardRead)
                {
                    args.State = CoreWebView2PermissionState.Allow;
                }
            };
            View.CoreWebView2.WebMessageReceived += OnMessage;
            var origin = new Uri(_url).GetLeftPart(UriPartial.Authority);
            View.CoreWebView2.NavigationStarting += (_, args) =>
            {
                if (args.Uri.StartsWith(origin, StringComparison.OrdinalIgnoreCase) ||
                    args.Uri.StartsWith("about:", StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                args.Cancel = true;
            };
            View.Source = new Uri(_url);
        }
        catch (Exception ex)
        {
            RuntimeMissing = true;
            MessageBox.Show(this, "需要安装 Microsoft Edge WebView2 运行时。\n" + ex.Message, "ExcelMergeFork");
            Close();
        }
    }

    private void OnMessage(object? sender, CoreWebView2WebMessageReceivedEventArgs e)
    {
        string text;
        try
        {
            text = e.TryGetWebMessageAsString();
        }
        catch
        {
            return;
        }

        if (text == "close")
        {
            Close();
            return;
        }

        if (text == "pick-folder")
        {
            PickFolder();
            return;
        }

        if (text.StartsWith("size:", StringComparison.Ordinal))
        {
            ApplySize(text["size:".Length..]);
        }
    }

    private void PickFolder()
    {
        var dialog = new OpenFolderDialog { Title = "选择备份根目录" };
        if (dialog.ShowDialog() != true || View.CoreWebView2 is null)
        {
            return;
        }

        View.CoreWebView2.PostWebMessageAsJson(JsonSerializer.Serialize(new { type = "folder", path = dialog.FolderName }));
    }

    private void ApplySize(string screen)
    {
        var (width, height, minWidth, minHeight, title) = screen switch
        {
            "merge" => (1520d, 820d, 1180d, 640d, "Excel 三向合并"),
            "compare" => (1280d, 840d, 1040d, 640d, "Excel 二向对比"),
            "choose" => (820d, 560d, 680d, 460d, "选择运行模式"),
            "backup" => (820d, 520d, 640d, 420d, "快速备份"),
            _ => (980d, 720d, 820d, 600d, "ExcelMergeFork 设置中心 v" + AppVersion.Display),
        };
        MinWidth = minWidth;
        MinHeight = minHeight;
        Width = Math.Max(width, minWidth);
        Height = Math.Max(height, minHeight);
        Title = title;
    }
}
