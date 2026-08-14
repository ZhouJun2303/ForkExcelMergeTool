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
