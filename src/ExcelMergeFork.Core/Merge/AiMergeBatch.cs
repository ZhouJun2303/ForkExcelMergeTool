using System.Security.Cryptography;
using System.Text;
using ExcelMergeFork.Core.Excel;
using ExcelMergeFork.Core.Git;

namespace ExcelMergeFork.Core.Merge;

public static class AiMergeBatch
{
    public static string StageRootFor(string repoRoot)
    {
        var hash = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(Path.GetFullPath(repoRoot))))[..16];
        return Path.Combine(Path.GetTempPath(), "ExcelMergeFork", "ai-stages", hash);
    }

    public static AiPromptBook MaterializePaths(string repoRoot, RepoConflictExcel file, string stageRoot) =>
        Materialize(repoRoot, file, stageRoot, withPreview: false);

    public static AiPromptBook MaterializeWithPreview(string repoRoot, RepoConflictExcel file, string stageRoot) =>
        Materialize(repoRoot, file, stageRoot, withPreview: true);

    private static AiPromptBook Materialize(string repoRoot, RepoConflictExcel file, string stageRoot, bool withPreview)
    {
        try
        {
            var directory = Path.Combine(stageRoot, SanitizeRelative(file.RelativePath));
            Directory.CreateDirectory(directory);
            var extension = Path.GetExtension(file.RelativePath);
            if (string.IsNullOrEmpty(extension))
            {
                extension = ".xlsx";
            }

            var local = Save(repoRoot, file.LocalBlob, Path.Combine(directory, "local" + extension));
            var basePath = Save(repoRoot, file.BaseBlob, Path.Combine(directory, "base" + extension));
            var remote = Save(repoRoot, file.RemoteBlob, Path.Combine(directory, "remote" + extension));
            MergePreview? preview = null;
            string? error = null;
            if (withPreview)
            {
                if (local.Length == 0 || remote.Length == 0)
                {
                    error = "缺少本地或线上文件，无法做三向冲突检测";
                }
                else if (basePath.Length == 0)
                {
                    error = "缺少基准文件，无法做三向冲突检测";
                }
                else
                {
                    try
                    {
                        using var session = new MergeSession(local, basePath, remote);
                        preview = PreviewBuilder.Build(session, new MergeOptions
                        {
                            AddNewSheets = false,
                            ResolveConflicts = true,
                        });
                    }
                    catch (Exception ex)
                    {
                        error = ex.Message;
                    }
                }
            }

            return new AiPromptBook
            {
                Label = file.RelativePath,
                WorktreePath = file.WorktreePath,
                Local = local,
                Base = basePath,
                Remote = remote,
                ReadError = error,
                Preview = preview,
            };
        }
        catch (Exception ex)
        {
            return new AiPromptBook
            {
                Label = file.RelativePath,
                WorktreePath = file.WorktreePath,
                ReadError = ex.Message,
            };
        }
    }

    private static string Save(string repoRoot, string? blob, string destination)
    {
        if (string.IsNullOrEmpty(blob))
        {
            return "";
        }

        GitBlobStore.WriteBlob(repoRoot, blob, destination);
        return destination;
    }

    private static string SanitizeRelative(string relative)
    {
        var invalid = Path.GetInvalidFileNameChars();
        var parts = relative.Split('/', StringSplitOptions.RemoveEmptyEntries);
        return string.Join(Path.DirectorySeparatorChar, parts.Select(part => SanitizeSegment(part, invalid)));
    }

    private static string SanitizeSegment(string segment, char[] invalid)
    {
        var text = new string(segment.Select(ch => invalid.Contains(ch) ? '_' : ch).ToArray()).Trim();
        return text.Length == 0 ? "_" : text;
    }
}
