using System.Windows;
using Wpf.Ui.Controls;
using UiMessageBox = Wpf.Ui.Controls.MessageBox;
using UiMessageBoxResult = Wpf.Ui.Controls.MessageBoxResult;

namespace ExcelMergeFork.App.Ui;

public static class FluentPrompt
{
    public static Task ShowInfoAsync(Window? owner, string message, string title) =>
        ShowAsync(owner, message, title, "确定", null);

    public static async Task<bool> ConfirmAsync(Window? owner, string message, string title)
    {
        var result = await ShowAsync(owner, message, title, "确定", "取消");
        return result == UiMessageBoxResult.Primary;
    }

    private static async Task<UiMessageBoxResult> ShowAsync(
        Window? owner,
        string message,
        string title,
        string primary,
        string? close)
    {
        var box = new UiMessageBox
        {
            Title = title,
            Content = new System.Windows.Controls.TextBlock
            {
                Text = message,
                TextWrapping = TextWrapping.Wrap,
                TextAlignment = TextAlignment.Left,
                MaxWidth = 420,
            },
            PrimaryButtonText = primary,
            PrimaryButtonAppearance = ControlAppearance.Primary,
            CloseButtonText = close ?? "",
            CloseButtonAppearance = ControlAppearance.Secondary,
            IsCloseButtonEnabled = close is not null,
            WindowStartupLocation = owner is null
                ? WindowStartupLocation.CenterScreen
                : WindowStartupLocation.CenterOwner,
            ShowTitle = true,
        };
        if (owner is not null)
        {
            box.Owner = owner;
        }

        return await box.ShowDialogAsync(true, CancellationToken.None);
    }
}
