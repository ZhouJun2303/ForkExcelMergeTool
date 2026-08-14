using System.Windows;
using ExcelMergeFork.App.ViewModels;

namespace ExcelMergeFork.App.Views;

public partial class DiffWindow
{
    private readonly DiffViewModel _vm;

    public DiffWindow(string local, string remote)
    {
        InitializeComponent();
        _vm = new DiffViewModel(local, remote);
        DataContext = _vm;
        Loaded += async (_, _) => await _vm.LoadAsync();
    }

    private void OpenSettings(object sender, RoutedEventArgs e) => App.OpenTracked(new SettingsWindow());

    private void OnExport(object sender, RoutedEventArgs e) => _vm.ExportCommand.Execute(false);

    private void OnClose(object sender, RoutedEventArgs e) => Close();
}
