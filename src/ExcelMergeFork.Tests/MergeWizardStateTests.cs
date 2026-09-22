using ExcelMergeFork.Core.Merge;

namespace ExcelMergeFork.Tests;

public class MergeWizardStateTests
{
    [Fact]
    public void EverySheetRequiresExplicitCompletion_AndNextRequiresCompletion()
    {
        var state = new MergeWizardState();
        state.ApplyPreview(Preview(
            new MergeChoice { Sheet = "Main", Key = "id-1", Kind = ConflictKind.Row },
            new PreviewItem { Sheet = "Main", Key = "id-1 (修改冲突)", Action = "将保留本地", Tag = PreviewTag.Conflict, ConflictIndex = 0 },
            new PreviewItem { Sheet = "Info", Key = "说明", Action = "信息", Tag = PreviewTag.Info }),
            ["Main", "Info"]);

        Assert.False(state.CanGenerate);
        Assert.False(state.Sheets[1].IsCompleted);
        Assert.False(state.MoveNext());

        Assert.True(state.TryCompleteCurrentSheet());
        Assert.True(state.MoveNext());
        Assert.Equal("Info", state.CurrentSheet?.Name);
        Assert.False(state.CanGenerate);
        Assert.True(state.TryCompleteCurrentSheet());
        Assert.True(state.CanGenerate);
    }

    [Fact]
    public void ChoicesAreRetainedAcrossSheets_AndAffectedCompletionResets()
    {
        var first = Preview(
            new MergeChoice { Sheet = "Main", Key = "id-1", Kind = ConflictKind.Row },
            new PreviewItem { Sheet = "Main", Key = "id-1 (修改冲突)", Action = "将保留本地", Tag = PreviewTag.Conflict, ConflictIndex = 0 });
        first = WithChoice(first, new MergeChoice { Sheet = "Extra", Key = "Code", Kind = ConflictKind.Column });
        first = AddItem(first, new PreviewItem { Sheet = "Extra", Key = "Code (列冲突)", Action = "将保留本地列", Tag = PreviewTag.Conflict, ConflictIndex = 1 });

        var state = new MergeWizardState();
        state.ApplyPreview(first, ["Main", "Extra"]);
        Assert.True(state.TrySetChoice("Main", "id-1", ConflictKind.Row, "remote"));
        Assert.True(state.TryCompleteCurrentSheet());
        Assert.True(state.MoveNext());
        Assert.True(state.TrySetChoice("Extra", "Code", ConflictKind.Column, "remote"));
        Assert.True(state.TryCompleteCurrentSheet());

        state.ApplyPreview(first, ["Main", "Extra"]);
        Assert.Equal("remote", state.Choices.Single(choice => choice.Sheet == "Main").Choice);
        Assert.Equal("remote", state.Choices.Single(choice => choice.Sheet == "Extra").Choice);
        Assert.True(state.CanGenerate);

        var changed = Preview(
            new MergeChoice { Sheet = "Main", Key = "id-1", Kind = ConflictKind.Row },
            new MergeChoice { Sheet = "Main", Key = "id-2", Kind = ConflictKind.Row },
            new PreviewItem { Sheet = "Main", Key = "id-1 (修改冲突)", Action = "将保留本地", Tag = PreviewTag.Conflict, ConflictIndex = 0 },
            new PreviewItem { Sheet = "Main", Key = "id-2 (修改冲突)", Action = "将保留本地", Tag = PreviewTag.Conflict, ConflictIndex = 1 });
        state.ApplyPreview(changed, ["Main", "Extra"]);

        Assert.False(state.Sheets.Single(sheet => sheet.Name == "Main").IsCompleted);
        Assert.False(state.Sheets.Single(sheet => sheet.Name == "Extra").IsCompleted);
        Assert.DoesNotContain(state.Choices, choice => choice.Sheet == "Extra");
        Assert.Equal(2, state.Choices.Count(choice => choice.Sheet == "Main"));
        Assert.Equal("remote", state.Choices.Single(choice => choice.Key == "id-1").Choice);
    }

    private static MergePreview Preview(params object[] parts)
    {
        var choices = parts.OfType<MergeChoice>().ToList();
        var items = parts.OfType<PreviewItem>().ToList();
        return new MergePreview
        {
            Items = items,
            ConflictEntries = choices,
            Summary = new PreviewSummary(),
            BaseSide = "local",
            Options = [],
        };
    }

    private static MergePreview WithChoice(MergePreview preview, MergeChoice choice) => new()
    {
        Items = preview.Items,
        ConflictEntries = preview.ConflictEntries.Concat([choice]).ToList(),
        Summary = preview.Summary,
        BaseSide = preview.BaseSide,
        Options = preview.Options,
    };

    private static MergePreview AddItem(MergePreview preview, PreviewItem item) => new()
    {
        Items = preview.Items.Concat([item]).ToList(),
        ConflictEntries = preview.ConflictEntries,
        Summary = preview.Summary,
        BaseSide = preview.BaseSide,
        Options = preview.Options,
    };
}
