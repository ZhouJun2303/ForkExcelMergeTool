using ExcelMergeFork.Core.Merge;

namespace ExcelMergeFork.Tests;

public class MergeAiPromptBuilderTests
{
    [Fact]
    public void PathPromptContainsAbsoluteWorkbookPathsWithoutConflictDetails()
    {
        var prompt = MergeAiPromptBuilder.BuildPathPrompt(
            "TestData/local.xlsx",
            "TestData/base.xlsx",
            "TestData/remote.xlsx",
            "TestData/merged.xlsx");

        Assert.Contains(Path.GetFullPath("TestData/local.xlsx"), prompt, StringComparison.Ordinal);
        Assert.Contains(Path.GetFullPath("TestData/base.xlsx"), prompt, StringComparison.Ordinal);
        Assert.Contains(Path.GetFullPath("TestData/remote.xlsx"), prompt, StringComparison.Ordinal);
        Assert.Contains(Path.GetFullPath("TestData/merged.xlsx"), prompt, StringComparison.Ordinal);
        Assert.DoesNotContain("Orders", prompt, StringComparison.Ordinal);
        Assert.DoesNotContain("order-42", prompt, StringComparison.Ordinal);
    }

    [Fact]
    public void DetailPromptContainsEverySheetConflictAndThreeSides()
    {
        var preview = new MergePreview
        {
            Items =
            [
                new PreviewItem
                {
                    Sheet = "Orders",
                    Key = "order-42 (修改冲突)",
                    Action = "将保留本地",
                    Tag = PreviewTag.Conflict,
                    ConflictIndex = 0,
                    BaseValues = ["Key: order-42", "Amount: 10"],
                    LocalValues = ["Key: order-42", "Amount: 11"],
                    RemoteValues = ["Key: order-42", "Amount: 12"],
                },
                new PreviewItem
                {
                    Sheet = "Summary",
                    Key = "Amount (列冲突)",
                    Action = "将保留线上列",
                    Tag = PreviewTag.Conflict,
                    ConflictIndex = 1,
                    BaseValues = ["Amount: 10"],
                    LocalValues = ["Amount: 11"],
                    RemoteValues = ["Amount: 12"],
                },
            ],
            ConflictEntries =
            [
                new MergeChoice { Sheet = "Orders", Key = "order-42", Kind = ConflictKind.Row, Choice = "local" },
                new MergeChoice { Sheet = "Summary", Key = "Amount", Kind = ConflictKind.Column, Choice = "remote" },
            ],
            Summary = new PreviewSummary { Conflict = 2 },
            BaseSide = "local",
            Options = ["G"],
        };

        var prompt = MergeAiPromptBuilder.BuildDetailPrompt(
            "local.xlsx", "base.xlsx", "remote.xlsx", "merged.xlsx", preview);

        Assert.Contains("Orders", prompt, StringComparison.Ordinal);
        Assert.Contains("Summary", prompt, StringComparison.Ordinal);
        Assert.Contains("order-42", prompt, StringComparison.Ordinal);
        Assert.Contains("Amount", prompt, StringComparison.Ordinal);
        Assert.Contains("行冲突", prompt, StringComparison.Ordinal);
        Assert.Contains("列冲突", prompt, StringComparison.Ordinal);
        Assert.Contains("Amount: 10", prompt, StringComparison.Ordinal);
        Assert.Contains("Amount: 11", prompt, StringComparison.Ordinal);
        Assert.Contains("Amount: 12", prompt, StringComparison.Ordinal);
        Assert.Contains("当前处理建议", prompt, StringComparison.Ordinal);
        Assert.Contains(Path.GetFullPath("local.xlsx"), prompt, StringComparison.Ordinal);
        Assert.Contains(Path.GetFullPath("base.xlsx"), prompt, StringComparison.Ordinal);
        Assert.Contains(Path.GetFullPath("remote.xlsx"), prompt, StringComparison.Ordinal);
        Assert.Contains(Path.GetFullPath("merged.xlsx"), prompt, StringComparison.Ordinal);
    }

}
