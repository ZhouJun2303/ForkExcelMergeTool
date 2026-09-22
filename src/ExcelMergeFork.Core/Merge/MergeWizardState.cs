namespace ExcelMergeFork.Core.Merge;

/// <summary>
/// Tracks completion and conflict choices for the Sheet-by-Sheet merge wizard.
/// This state is independent from the currently visible preview rows so choices
/// remain available when the user navigates between Sheets.
/// </summary>
public sealed class MergeWizardState
{
    private readonly List<MergeSheetState> _sheets = [];
    private readonly Dictionary<string, MergeChoice> _choices = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _signatures = new(StringComparer.Ordinal);
    private int _currentIndex;

    public IReadOnlyList<MergeSheetState> Sheets => _sheets;
    public int CurrentIndex => _currentIndex;
    public MergeSheetState? CurrentSheet =>
        _currentIndex >= 0 && _currentIndex < _sheets.Count ? _sheets[_currentIndex] : null;
    public bool CanGenerate => _sheets.Count > 0 && _sheets.All(sheet => sheet.IsCompleted);
    public IReadOnlyList<MergeChoice> Choices => _choices.Values.Select(CloneChoice).ToList();

    public void ApplyPreview(MergePreview preview, IReadOnlyList<string> orderedSheetNames)
    {
        var previousCurrent = CurrentSheet?.Name;
        var previousSheets = _sheets.ToDictionary(sheet => sheet.Name, StringComparer.Ordinal);
        var previousSignatures = new Dictionary<string, string>(_signatures, StringComparer.Ordinal);
        var nextChoices = new Dictionary<string, MergeChoice>(StringComparer.Ordinal);
        foreach (var incoming in preview.ConflictEntries)
        {
            nextChoices[incoming.ChoiceKey] = _choices.TryGetValue(incoming.ChoiceKey, out var saved)
                ? CloneChoice(saved)
                : CloneChoice(incoming);
        }

        _choices.Clear();
        foreach (var pair in nextChoices)
        {
            _choices[pair.Key] = pair.Value;
        }

        _sheets.Clear();
        _signatures.Clear();
        foreach (var sheetName in orderedSheetNames.Distinct(StringComparer.Ordinal))
        {
            var selectable = preview.ConflictEntries
                .Where(choice => string.Equals(choice.Sheet, sheetName, StringComparison.Ordinal))
                .Select(choice => choice.ChoiceKey)
                .ToHashSet(StringComparer.Ordinal);
            var signature = BuildSignature(sheetName, preview, selectable);
            var hasSelectable = selectable.Count > 0;
            var state = new MergeSheetState
            {
                Name = sheetName,
                HasSelectableConflicts = hasSelectable,
                IsCompleted = false,
            };

            if (previousSheets.TryGetValue(sheetName, out var previous) &&
                previousSignatures.TryGetValue(sheetName, out var previousSignature) &&
                string.Equals(previousSignature, signature, StringComparison.Ordinal))
            {
                state.IsCompleted = previous.IsCompleted;
            }

            _sheets.Add(state);
            _signatures[sheetName] = signature;
        }

        if (_sheets.Count == 0)
        {
            _currentIndex = 0;
            return;
        }

        var preservedIndex = previousCurrent is null
            ? -1
            : _sheets.FindIndex(sheet => string.Equals(sheet.Name, previousCurrent, StringComparison.Ordinal));
        _currentIndex = preservedIndex >= 0
            ? preservedIndex
            : Math.Clamp(_currentIndex, 0, _sheets.Count - 1);
    }

    public bool TrySetChoice(string sheet, string key, ConflictKind kind, string choice)
    {
        if (!string.Equals(choice, "local", StringComparison.Ordinal) &&
            !string.Equals(choice, "remote", StringComparison.Ordinal))
        {
            return false;
        }

        var candidate = new MergeChoice { Sheet = sheet, Key = key, Kind = kind, Choice = choice };
        if (!_choices.ContainsKey(candidate.ChoiceKey))
        {
            return false;
        }

        _choices[candidate.ChoiceKey].Choice = choice;
        var sheetState = _sheets.FirstOrDefault(item =>
            string.Equals(item.Name, sheet, StringComparison.Ordinal));
        if (sheetState is not null && sheetState.HasSelectableConflicts)
        {
            sheetState.IsCompleted = false;
        }
        return true;
    }

    public bool TryCompleteCurrentSheet()
    {
        var current = CurrentSheet;
        if (current is null)
        {
            return false;
        }

        current.IsCompleted = true;
        return true;
    }

    public bool MovePrevious()
    {
        if (_currentIndex <= 0 || _sheets.Count == 0)
        {
            return false;
        }

        _currentIndex--;
        return true;
    }

    public bool MoveNext()
    {
        if (_currentIndex >= _sheets.Count - 1 || CurrentSheet is null || !CurrentSheet.IsCompleted)
        {
            return false;
        }

        _currentIndex++;
        return true;
    }

    private static string BuildSignature(
        string sheetName,
        MergePreview preview,
        IReadOnlySet<string> selectable)
    {
        var itemParts = preview.Items
            .Where(item => string.Equals(item.Sheet, sheetName, StringComparison.Ordinal))
            .Select(item =>
            {
                var conflict = item.ConflictIndex is int index &&
                               index >= 0 &&
                               index < preview.ConflictEntries.Count
                    ? preview.ConflictEntries[index].ChoiceKey
                    : "";
                return string.Join("\u0001", item.Key, item.Action, item.Tag, conflict);
            });
        var conflictParts = selectable.OrderBy(key => key, StringComparer.Ordinal);
        return string.Join("\u0002", itemParts.Concat(conflictParts));
    }

    private static MergeChoice CloneChoice(MergeChoice choice) => new()
    {
        Sheet = choice.Sheet,
        Key = choice.Key,
        Choice = choice.Choice,
        Kind = choice.Kind,
        AutoType = choice.AutoType,
    };
}

public sealed class MergeSheetState
{
    public required string Name { get; init; }
    public bool HasSelectableConflicts { get; internal set; }
    public bool IsCompleted { get; internal set; }
}
