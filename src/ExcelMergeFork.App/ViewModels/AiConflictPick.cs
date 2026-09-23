using CommunityToolkit.Mvvm.ComponentModel;
using ExcelMergeFork.Core.Git;

namespace ExcelMergeFork.App.ViewModels;

public sealed partial class AiConflictPick : ObservableObject
{
    public required string RelativePath { get; init; }
    public required string FullPath { get; init; }
    public required string StatusLabel { get; init; }
    public string SidesText { get; init; } = "";
    public string ExtraPathText { get; init; } = "";
    public bool IsCurrent { get; init; }
    public bool HasBothSides { get; init; }
    public bool UseLaunchPaths { get; init; }
    public string? RepoRoot { get; init; }
    public RepoConflictExcel? Source { get; init; }
    [ObservableProperty] private bool _isChecked;

    public event EventHandler? CheckedChanged;

    partial void OnIsCheckedChanged(bool value) => CheckedChanged?.Invoke(this, EventArgs.Empty);
}
