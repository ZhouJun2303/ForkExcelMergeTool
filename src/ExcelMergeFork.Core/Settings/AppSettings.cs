using System.Text.Json;
using System.Text.Json.Nodes;

namespace ExcelMergeFork.Core.Settings;

public static class StartupFeature
{
    public const string BackupOnly = "backup_only";
    public const string MergeDiff = "merge_diff";
    public const string AskEachTime = "ask_each_time";
}

public sealed class UserSettings
{
    public bool SkipNewRows { get; set; }
    public bool SkipNewColumns { get; set; }
    public bool DeleteMissingRows { get; set; }
    public bool DeleteMissingColumns { get; set; }
    public bool AddNewSheets { get; set; } = true;
    public bool DeleteMissingSheets { get; set; }
    public bool ResolveConflicts { get; set; } = true;
    public string StartupFeatureValue { get; set; } = StartupFeature.MergeDiff;
    public string BackupRootDir { get; set; } = "";
    public bool AutoOpenMerged { get; set; } = true;
    public bool AutoOpenCompare { get; set; }
    public bool DarkTheme { get; set; }
    public Dictionary<string, bool> DiffFilter { get; set; } = new()
    {
        ["新增行"] = true,
        ["删除行"] = true,
        ["新增列"] = true,
        ["删除列"] = true,
        ["修改"] = true,
    };
}

public static class AppSettingsStore
{
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = null,
    };

    public static UserSettings Load()
    {
        var settings = new UserSettings();
        var data = ReadRaw();
        if (data is null)
        {
            return settings;
        }

        settings.SkipNewRows = Bool(data, "A", settings.SkipNewRows);
        settings.SkipNewColumns = Bool(data, "B", settings.SkipNewColumns);
        settings.DeleteMissingRows = Bool(data, "C", settings.DeleteMissingRows);
        settings.DeleteMissingColumns = Bool(data, "D", settings.DeleteMissingColumns);
        settings.AddNewSheets = Bool(data, "E", settings.AddNewSheets);
        settings.DeleteMissingSheets = Bool(data, "F", settings.DeleteMissingSheets);
        settings.ResolveConflicts = Bool(data, "G", settings.ResolveConflicts);
        settings.StartupFeatureValue = data["startup_feature"]?.GetValue<string>() ?? settings.StartupFeatureValue;
        settings.BackupRootDir = data["backup_root_dir"]?.GetValue<string>() ?? "";
        settings.AutoOpenMerged = Bool(data, "auto_open_merged", true);
        settings.AutoOpenCompare = Bool(data, "auto_open_compare", false);
        settings.DarkTheme = Bool(data, "dark_theme", false);
        if (data["diff_filter"] is JsonObject filter)
        {
            foreach (var kv in filter)
            {
                if (kv.Value is JsonValue value && value.TryGetValue<bool>(out var flag))
                {
                    settings.DiffFilter[kv.Key] = flag;
                }
            }
        }

        return settings;
    }

    public static void Save(UserSettings settings)
    {
        var data = ReadRaw() ?? new JsonObject();
        data["A"] = settings.SkipNewRows;
        data["B"] = settings.SkipNewColumns;
        data["C"] = settings.DeleteMissingRows;
        data["D"] = settings.DeleteMissingColumns;
        data["E"] = settings.AddNewSheets;
        data["F"] = settings.DeleteMissingSheets;
        data["G"] = settings.ResolveConflicts;
        data["merge_options_schema"] = AppConstants.MergeOptionsSchema;
        data["startup_feature"] = settings.StartupFeatureValue;
        if (string.IsNullOrWhiteSpace(settings.BackupRootDir))
        {
            data.Remove("backup_root_dir");
        }
        else
        {
            data["backup_root_dir"] = Path.GetFullPath(settings.BackupRootDir);
        }

        data["auto_open_merged"] = settings.AutoOpenMerged;
        data["auto_open_compare"] = settings.AutoOpenCompare;
        data["dark_theme"] = settings.DarkTheme;
        var filter = new JsonObject();
        foreach (var kv in settings.DiffFilter)
        {
            filter[kv.Key] = kv.Value;
        }

        data["diff_filter"] = filter;
        Directory.CreateDirectory(Path.GetDirectoryName(AppPaths.OptionsFile) ?? ".");
        File.WriteAllText(AppPaths.OptionsFile, data.ToJsonString(JsonOptions));
    }

    private static JsonObject? ReadRaw()
    {
        try
        {
            if (!File.Exists(AppPaths.OptionsFile))
            {
                return null;
            }

            return JsonNode.Parse(File.ReadAllText(AppPaths.OptionsFile)) as JsonObject;
        }
        catch
        {
            return null;
        }
    }

    private static bool Bool(JsonObject data, string key, bool fallback)
    {
        if (data[key] is JsonValue value && value.TryGetValue<bool>(out var flag))
        {
            return flag;
        }

        return fallback;
    }
}
