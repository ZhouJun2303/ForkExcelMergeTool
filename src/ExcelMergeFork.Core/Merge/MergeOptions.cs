namespace ExcelMergeFork.Core.Merge;

public sealed class MergeOptions
{
    public bool SkipNewRows { get; set; }
    public bool SkipNewColumns { get; set; }
    public bool DeleteMissingRows { get; set; }
    public bool DeleteMissingColumns { get; set; }
    public bool AddNewSheets { get; set; } = true;
    public bool DeleteMissingSheets { get; set; }
    public bool ResolveConflicts { get; set; } = true;
    public string BaseSide { get; set; } = "local";

    public static MergeOptions FromLetters(IEnumerable<string> letters, string baseSide = "local")
    {
        var set = new HashSet<string>(letters.Select(x => x.ToUpperInvariant()), StringComparer.OrdinalIgnoreCase);
        return new MergeOptions
        {
            SkipNewRows = set.Contains("A"),
            SkipNewColumns = set.Contains("B"),
            DeleteMissingRows = set.Contains("C"),
            DeleteMissingColumns = set.Contains("D"),
            AddNewSheets = set.Contains("E"),
            DeleteMissingSheets = set.Contains("F"),
            ResolveConflicts = set.Contains("G"),
            BaseSide = baseSide,
        };
    }

    public static MergeOptions ForMode(string mode, string baseSide = "local")
    {
        return mode.ToUpperInvariant() switch
        {
            "A" => new MergeOptions
            {
                SkipNewRows = false,
                SkipNewColumns = true,
                DeleteMissingRows = false,
                DeleteMissingColumns = false,
                AddNewSheets = false,
                DeleteMissingSheets = false,
                ResolveConflicts = false,
                BaseSide = baseSide,
            },
            "B" => new MergeOptions
            {
                SkipNewRows = true,
                SkipNewColumns = false,
                DeleteMissingRows = false,
                DeleteMissingColumns = false,
                AddNewSheets = false,
                DeleteMissingSheets = false,
                ResolveConflicts = false,
                BaseSide = baseSide,
            },
            "C" => new MergeOptions
            {
                SkipNewRows = true,
                SkipNewColumns = true,
                DeleteMissingRows = false,
                DeleteMissingColumns = false,
                AddNewSheets = true,
                DeleteMissingSheets = false,
                ResolveConflicts = false,
                BaseSide = baseSide,
            },
            "D" => new MergeOptions
            {
                SkipNewRows = true,
                SkipNewColumns = true,
                DeleteMissingRows = false,
                DeleteMissingColumns = false,
                AddNewSheets = false,
                DeleteMissingSheets = false,
                ResolveConflicts = true,
                BaseSide = baseSide,
            },
            _ => FromLetters(["E", "G"], baseSide),
        };
    }

    public HashSet<string> ToLetterSet()
    {
        var set = new HashSet<string>(StringComparer.Ordinal);
        if (SkipNewRows) set.Add("A");
        if (SkipNewColumns) set.Add("B");
        if (DeleteMissingRows) set.Add("C");
        if (DeleteMissingColumns) set.Add("D");
        if (AddNewSheets) set.Add("E");
        if (DeleteMissingSheets) set.Add("F");
        if (ResolveConflicts) set.Add("G");
        return set;
    }
}
