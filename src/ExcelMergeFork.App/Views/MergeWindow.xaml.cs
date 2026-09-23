using System.Windows;
using ExcelMergeFork.App.ViewModels;
using ExcelMergeFork.Core.Git;

namespace ExcelMergeFork.App.Views;

public partial class MergeWindow
{
    private readonly MergeViewModel _vm;
    private readonly bool _gitDriver;

    public bool WriteBackSucceeded { get; private set; }

    public MergeWindow(string local, string basePath, string remote, string merged, GitDriverRequest? driver = null)
    {
        InitializeComponent();
        _gitDriver = driver is not null;
        _vm = new MergeViewModel(local, basePath, remote, merged, driver);
        DataContext = _vm;
        Title = "Excel 三向合并";
        Loaded += async (_, _) => await _vm.LoadAsync();
        Closed += (_, _) => _vm.DisposeSession();
    }

    private void OpenSettings(object sender, RoutedEventArgs e) => App.OpenSettings();

    private void OnCancel(object sender, RoutedEventArgs e) => Close();

    private void ToggleAiPanel(object sender, RoutedEventArgs e)
    {
        var open = AiPanel.Visibility != Visibility.Visible;
        AiPanel.Visibility = open ? Visibility.Visible : Visibility.Collapsed;
        AiToggleButton.Content = open ? "收起 AI 合并" : "AI 合并";
    }

    private async void CopySelectedPathsToAi(object sender, RoutedEventArgs e) =>
        await CopyCheckedPromptAsync(detail: false);

    private async void CopySelectedConflictsToAi(object sender, RoutedEventArgs e) =>
        await CopyCheckedPromptAsync(detail: true);

    private async Task CopyCheckedPromptAsync(bool detail)
    {
        try
        {
            var prompt = detail
                ? await _vm.BuildCheckedDetailPromptAsync()
                : await _vm.BuildCheckedPathPromptAsync();
            if (string.IsNullOrWhiteSpace(prompt))
            {
                return;
            }

            CopyAiPrompt(prompt);
        }
        catch (Exception ex)
        {
            _vm.AiStatusText = "复制失败：" + ex.Message;
        }
    }

    private void CopyAiPrompt(string prompt)
    {
        try
        {
            Clipboard.SetText(prompt);
            Close();
        }
        catch (Exception ex)
        {
            _vm.StatusText = "复制失败：" + ex.Message;
            _vm.AiStatusText = "复制失败：" + ex.Message;
        }
    }

    private async void OnGenerate(object sender, RoutedEventArgs e)
    {
        if (!_vm.AllSheetsConfirmed)
        {
            var pending = _vm.UnconfirmedSheetNames;
            var detail = pending.Count == 0
                ? "没有可确认的 Sheet。"
                : "以下 Sheet 尚未确认：\n" + string.Join("\n", pending);
            _vm.StatusText = "请先确认全部 Sheet。";
            MessageBox.Show(
                this,
                detail + "\n\n请先对每个 Sheet 点击“确认本 Sheet”，再生成合并结果。",
                "ExcelMergeFork",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return;
        }

        await _vm.GenerateCommand.ExecuteAsync(null);
    }

    private void OnConfirm(object sender, RoutedEventArgs e)
    {
        var result = _vm.Confirm();
        if (result.Success)
        {
            if (_gitDriver)
            {
                WriteBackSucceeded = true;
            }

            MessageBox.Show(string.IsNullOrWhiteSpace(result.Message) ? "冲突已解决。" : result.Message, "ExcelMergeFork");
            Close();
            return;
        }

        MessageBox.Show(string.Join(Environment.NewLine, result.Errors.DefaultIfEmpty("确认失败")), "ExcelMergeFork");
    }
}
