using ClosedXML.Excel;

namespace ExcelMergeFork.Core.Excel;

public sealed class WorkbookSession : IDisposable
{
    public string Path { get; }
    public XLWorkbook Workbook { get; }
    public IReadOnlyList<string> AllSheetNames { get; }
    public IReadOnlyList<string> SheetNames { get; }
    public Dictionary<string, SheetSnapshot> Sheets { get; }

    private WorkbookSession(string path, XLWorkbook workbook)
    {
        Path = path;
        Workbook = workbook;
        AllSheetNames = workbook.Worksheets.Select(ws => ws.Name).ToList();
        SheetNames = AllSheetNames.Where(name => !SheetFilter.ShouldSkip(name)).ToList();
        Sheets = new Dictionary<string, SheetSnapshot>(StringComparer.Ordinal);
        foreach (var name in SheetNames)
        {
            Sheets[name] = SheetSnapshot.From(workbook.Worksheet(name));
        }
    }

    public static WorkbookSession Open(string path)
    {
        if (!File.Exists(path))
        {
            throw new FileNotFoundException("Excel 文件不存在: " + path, path);
        }

        return new WorkbookSession(path, new XLWorkbook(path));
    }

    public bool HasSheet(string name) => Workbook.Worksheets.Contains(name);

    public IXLWorksheet? Worksheet(string name) =>
        Workbook.Worksheets.Contains(name) ? Workbook.Worksheet(name) : null;

    public SheetSnapshot Snapshot(string name) =>
        Sheets.TryGetValue(name, out var snapshot) ? snapshot : SheetSnapshot.Empty(name);

    public void RefreshSnapshot(string name)
    {
        var ws = Worksheet(name);
        if (ws is null)
        {
            Sheets.Remove(name);
            return;
        }

        Sheets[name] = SheetSnapshot.From(ws);
    }

    public void Dispose() => Workbook.Dispose();
}

public sealed class MergeSession : IDisposable
{
    public WorkbookSession Local { get; }
    public WorkbookSession Base { get; }
    public WorkbookSession Remote { get; }

    public MergeSession(string localPath, string basePath, string remotePath)
    {
        Local = WorkbookSession.Open(localPath);
        Base = WorkbookSession.Open(basePath);
        Remote = WorkbookSession.Open(remotePath);
    }

    public IReadOnlyList<string> UnionSheetNames()
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        var names = new List<string>();
        foreach (var source in new[] { Base.SheetNames, Local.SheetNames, Remote.SheetNames })
        {
            foreach (var name in source)
            {
                if (seen.Add(name))
                {
                    names.Add(name);
                }
            }
        }

        return names;
    }

    public void Dispose()
    {
        Local.Dispose();
        Base.Dispose();
        Remote.Dispose();
    }
}
