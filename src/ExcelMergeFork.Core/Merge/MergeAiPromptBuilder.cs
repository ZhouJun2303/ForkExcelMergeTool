using System.Text;
using ExcelMergeFork.Core.Excel;

namespace ExcelMergeFork.Core.Merge;

public static class MergeAiPromptBuilder
{
    public static string BuildPathsPrompt(string local, string basePath, string remote, string merged) =>
        BuildPathPrompt(local, basePath, remote, merged);

    public static string BuildPathPrompt(string local, string basePath, string remote, string merged)
    {
        var paths = NormalizePaths(local, basePath, remote, merged);
        var builder = new StringBuilder();
        builder.AppendLine("你是 Excel 三向合并助手。请读取下面四个工作簿的绝对路径，自行检查全部可处理 Sheet（名称以 # 开头的 Sheet 按规则跳过），识别所有行和列冲突，并给出逐 Sheet 的处理建议。不要修改文件。");
        AppendPaths(builder, paths);
        builder.AppendLine("请在建议中明确 Sheet、首列行键或列名、冲突类型，以及应该保留 LOCAL 还是 REMOTE，并说明依据。");
        return builder.ToString();
    }

    public static string BuildDetailPrompt(
        string local,
        string basePath,
        string remote,
        string merged,
        MergePreview preview)
    {
        var paths = NormalizePaths(local, basePath, remote, merged);
        var builder = new StringBuilder();
        builder.AppendLine("你是 Excel 三向合并助手。下面是工作簿路径和已经检测到的全部逐 Sheet 冲突。请逐项检查并给出处理建议，不要修改文件。");
        AppendPaths(builder, paths);
        builder.AppendLine();
        builder.AppendLine("冲突详情：");

        var conflicts = preview.Items
            .Where(item => item.ConflictIndex is int index && index >= 0 && index < preview.ConflictEntries.Count)
            .ToList();
        if (conflicts.Count == 0)
        {
            builder.AppendLine("（当前预览没有需要手工选择的冲突。）");
            return builder.ToString();
        }

        foreach (var item in conflicts)
        {
            var conflict = preview.ConflictEntries[item.ConflictIndex!.Value];
            builder.AppendLine($"- Sheet: {conflict.Sheet}");
            builder.AppendLine($"  冲突类型: {(conflict.Kind == ConflictKind.Column ? "列冲突" : "行冲突")}");
            builder.AppendLine($"  {(conflict.Kind == ConflictKind.Column ? "列名" : "行键")}: {conflict.Key}");
            builder.AppendLine($"  当前处理建议: {item.Action}");
            builder.AppendLine($"  当前选择: {(conflict.Choice == "remote" ? "REMOTE" : "LOCAL")}");
            builder.AppendLine("  BASE: " + FormatValues(item.BaseValues));
            builder.AppendLine("  LOCAL: " + FormatValues(item.LocalValues));
            builder.AppendLine("  REMOTE: " + FormatValues(item.RemoteValues));
        }

        return builder.ToString();
    }

    public static string BuildDetailPrompt(
        string local,
        string basePath,
        string remote,
        string merged,
        MergeSession session)
    {
        var preview = PreviewBuilder.Build(session, new MergeOptions
        {
            AddNewSheets = false,
            ResolveConflicts = true,
        });
        return BuildDetailPrompt(local, basePath, remote, merged, preview);
    }

    public static string BuildConflictPrompt(
        string local,
        string basePath,
        string remote,
        string merged,
        MergePreview preview) =>
        BuildDetailPrompt(local, basePath, remote, merged, preview);

    private static (string Local, string Base, string Remote, string Merged) NormalizePaths(
        string local,
        string basePath,
        string remote,
        string merged) =>
        (Path.GetFullPath(local), Path.GetFullPath(basePath), Path.GetFullPath(remote), Path.GetFullPath(merged));

    private static void AppendPaths(
        StringBuilder builder,
        (string Local, string Base, string Remote, string Merged) paths)
    {
        builder.AppendLine("LOCAL: " + paths.Local);
        builder.AppendLine("BASE: " + paths.Base);
        builder.AppendLine("REMOTE: " + paths.Remote);
        builder.AppendLine("MERGED: " + paths.Merged);
    }

    private static string FormatValues(IReadOnlyList<string> values) =>
        values.Count == 0 ? "（此侧没有该行或列）" : string.Join("；", values);
}
