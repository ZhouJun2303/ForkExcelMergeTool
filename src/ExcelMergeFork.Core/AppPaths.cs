namespace ExcelMergeFork.Core;

public static class AppPaths
{
    public static string Home
    {
        get
        {
            var env = Environment.GetEnvironmentVariable(AppConstants.HomeEnvVar);
            if (!string.IsNullOrWhiteSpace(env))
            {
                return Path.GetFullPath(env);
            }

            return AppContext.BaseDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }
    }

    public static string LogFile => Path.Combine(Home, AppConstants.LogFileName);

    public static string OptionsFile => Path.Combine(Home, AppConstants.OptionsFileName);
}
