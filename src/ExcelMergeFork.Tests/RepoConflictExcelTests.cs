using ExcelMergeFork.Core.Git;
using ExcelMergeFork.Core.Merge;

namespace ExcelMergeFork.Tests;

public class RepoConflictExcelTests
{
    private const string BaseBlob = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa";
    private const string LocalBlob = "bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb";
    private const string RemoteBlob = "cccccccccccccccccccccccccccccccccccccccc";

    [Fact]
    public void Parse_KeepsBothModifiedExcelAndSkipsOtherFiles()
    {
        var repo = Path.Combine(Path.GetTempPath(), "emf-parse-repo");
        var output = string.Join('\n',
            Record(BaseBlob, 1, "sheets/a.xlsx"),
            Record(LocalBlob, 2, "sheets/a.xlsx"),
            Record(RemoteBlob, 3, "sheets/a.xlsx"),
            Record("dddddddddddddddddddddddddddddddddddddddd", 2, "notes.txt"),
            Record("eeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeeee", 1, "legacy.xls"),
            Record("ffffffffffffffffffffffffffffffffffffffff", 2, "legacy.xls"),
            Record("1111111111111111111111111111111111111111", 3, "legacy.xls"),
            Record(LocalBlob, 2, "only-local.xlsx"),
            Record(BaseBlob, 1, "only-local.xlsx"),
            Record(BaseBlob, 1, "../secret.xlsx"),
            Record(LocalBlob, 2, "b.xlsx"),
            Record(RemoteBlob, 3, "b.xlsx"));

        var files = RepoConflictExcelFinder.Parse(repo, output);

        Assert.Equal(["b.xlsx", "only-local.xlsx", "sheets/a.xlsx"], files.Select(file => file.RelativePath).ToArray());
        var both = files.Single(file => file.RelativePath == "sheets/a.xlsx");
        Assert.True(both.HasBothSides);
        Assert.Equal(BaseBlob, both.BaseBlob);
        Assert.Equal(LocalBlob, both.LocalBlob);
        Assert.Equal(RemoteBlob, both.RemoteBlob);
        Assert.Equal(Path.Combine(Path.GetFullPath(repo), "sheets", "a.xlsx"), both.WorktreePath);
        var incomplete = files.Single(file => file.RelativePath == "only-local.xlsx");
        Assert.False(incomplete.HasBothSides);
        Assert.Null(incomplete.RemoteBlob);
    }

    [Fact]
    public void Parse_ReadsNulSeparatedRecords()
    {
        var output = Record(LocalBlob, 2, "dir/book.xlsx") + "\0" + Record(RemoteBlob, 3, "dir/book.xlsx") + "\0";
        var files = RepoConflictExcelFinder.Parse(Path.GetTempPath(), output);
        var file = Assert.Single(files);
        Assert.Equal("dir/book.xlsx", file.RelativePath);
        Assert.True(file.HasBothSides);
        Assert.Null(file.BaseBlob);
    }

    [Fact]
    public void BatchPrompts_ListCheckedPathsAndConflictOrError()
    {
        var books = new[]
        {
            new AiPromptBook
            {
                Label = "sheets/a.xlsx",
                WorktreePath = "D:/repo/sheets/a.xlsx",
                Local = "D:/stages/a/local.xlsx",
                Base = "D:/stages/a/base.xlsx",
                Remote = "D:/stages/a/remote.xlsx",
                Merged = "D:/repo/sheets/a.xlsx",
                Preview = new MergePreview
                {
                    Items =
                    [
                        new PreviewItem
                        {
                            Sheet = "Orders",
                            Key = "order-42 (修改冲突)",
                            Action = "将保留线上",
                            Tag = PreviewTag.Conflict,
                            ConflictIndex = 0,
                            BaseValues = ["Amount: 10"],
                            LocalValues = ["Amount: 11"],
                            RemoteValues = ["Amount: 12"],
                        },
                    ],
                    ConflictEntries =
                    [
                        new MergeChoice { Sheet = "Orders", Key = "order-42", Kind = ConflictKind.Row, Choice = "remote" },
                    ],
                    Summary = new PreviewSummary { Conflict = 1 },
                    BaseSide = "local",
                    Options = ["G"],
                },
            },
            new AiPromptBook
            {
                Label = "only.xlsx",
                WorktreePath = "D:/repo/only.xlsx",
                Local = "D:/stages/only/local.xlsx",
                Remote = "D:/stages/only/remote.xlsx",
                ReadError = "缺少基准文件，无法做三向冲突检测",
            },
        };

        var paths = MergeAiPromptBuilder.BuildBatchPathPrompt(books);
        Assert.Contains("sheets/a.xlsx", paths, StringComparison.Ordinal);
        Assert.Contains(Path.GetFullPath("D:/repo/sheets/a.xlsx"), paths, StringComparison.Ordinal);
        Assert.Contains(Path.GetFullPath("D:/stages/a/local.xlsx"), paths, StringComparison.Ordinal);
        Assert.Contains(Path.GetFullPath("D:/stages/only/remote.xlsx"), paths, StringComparison.Ordinal);
        Assert.Contains("（此侧没有该文件）", paths, StringComparison.Ordinal);
        Assert.Contains("（尚未生成，不要写入文件）", paths, StringComparison.Ordinal);
        Assert.Contains("缺少基准文件", paths, StringComparison.Ordinal);
        Assert.DoesNotContain("order-42", paths, StringComparison.Ordinal);

        var detail = MergeAiPromptBuilder.BuildBatchDetailPrompt(books);
        Assert.Contains("order-42", detail, StringComparison.Ordinal);
        Assert.Contains("REMOTE", detail, StringComparison.Ordinal);
        Assert.Contains("Amount: 12", detail, StringComparison.Ordinal);
        Assert.Contains("无法读取冲突：缺少基准文件", detail, StringComparison.Ordinal);
        Assert.Contains(Path.GetFullPath("D:/repo/only.xlsx"), detail, StringComparison.Ordinal);
    }

    [Fact]
    public void Scan_FindsConflictedXlsxAndExportsReadableStages()
    {
        var root = Directory.CreateTempSubdirectory("emf-conflict-").FullName;
        var stage = Directory.CreateTempSubdirectory("emf-stages-").FullName;
        try
        {
            Git(root, "init", "-b", "main");
            Directory.CreateDirectory(Path.Combine(root, ".git", "info"));
            File.WriteAllText(Path.Combine(root, ".git", "info", "attributes"), "*.xlsx binary\n*.xltx binary\n");
            var book = Path.Combine(root, "Data", "book.xlsx");
            Directory.CreateDirectory(Path.GetDirectoryName(book)!);
            File.Copy(TestRepo.Fixture("base.xlsx"), book);
            File.WriteAllText(Path.Combine(root, "notes.txt"), "base");
            File.Copy(TestRepo.Fixture("base.xlsx"), Path.Combine(root, "same.xlsx"));
            Git(root, "add", "--", "Data/book.xlsx", "notes.txt", "same.xlsx");
            Commit(root, "base");
            Git(root, "checkout", "-b", "other");
            File.Copy(TestRepo.Fixture("remote.xlsx"), book, true);
            File.WriteAllText(Path.Combine(root, "notes.txt"), "remote");
            Git(root, "add", "--", "Data/book.xlsx", "notes.txt");
            Commit(root, "remote");
            Git(root, "checkout", "main");
            File.Copy(TestRepo.Fixture("local.xlsx"), book, true);
            File.WriteAllText(Path.Combine(root, "notes.txt"), "local");
            Git(root, "add", "--", "Data/book.xlsx", "notes.txt");
            Commit(root, "local");

            var merge = GitRunner.Run(root, ["merge", "other"], 30);
            Assert.NotEqual(0, merge.ExitCode);

            var scan = RepoConflictExcelFinder.Scan(root);
            Assert.Null(scan.Error);
            var file = Assert.Single(scan.Files);
            Assert.Equal("Data/book.xlsx", file.RelativePath);
            Assert.True(file.HasBothSides);
            Assert.False(string.IsNullOrEmpty(file.BaseBlob));
            Assert.Equal(Path.GetFullPath(book), file.WorktreePath);

            var promptBook = AiMergeBatch.MaterializeWithPreview(root, file, stage);
            Assert.True(string.IsNullOrEmpty(promptBook.ReadError), promptBook.ReadError);
            Assert.True(File.Exists(promptBook.Local));
            Assert.True(File.Exists(promptBook.Base));
            Assert.True(File.Exists(promptBook.Remote));
            Assert.Equal(new byte[] { 0x50, 0x4B }, File.ReadAllBytes(promptBook.Local)[..2]);
            Assert.NotNull(promptBook.Preview);
            Assert.Contains("Data/book.xlsx", MergeAiPromptBuilder.BuildBatchPathPrompt([promptBook]), StringComparison.Ordinal);
        }
        finally
        {
            TryDelete(stage);
            TryDelete(root);
        }
    }

    private static string Record(string blob, int stage, string path) => $"100644 {blob} {stage}\t{path}";

    private static void Git(string repo, params string[] args)
    {
        var result = GitRunner.Run(repo, args, 30);
        Assert.True(result.ExitCode == 0, result.StdErr + result.StdOut);
    }

    private static void Commit(string repo, string message)
    {
        Git(repo, "-c", "user.email=test@example.com", "-c", "user.name=ExcelMergeFork", "commit", "--no-verify", "-m", message);
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, true);
            }
        }
        catch
        {
        }
    }
}
