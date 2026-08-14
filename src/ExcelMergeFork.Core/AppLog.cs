namespace ExcelMergeFork.Core;

public static class AppLog
{
    private static readonly object Gate = new();

    public static void Info(string message) => Write(message, error: false);

    public static void Error(string message) => Write(message, error: true);

    public static void Exception(string message, Exception ex)
    {
        Error(message + ": " + ex.Message);
        Error(ex.ToString());
    }

    private static void Write(string message, bool error)
    {
        try
        {
            var line = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {(error ? "[ERROR] " : "")}{message}{Environment.NewLine}";
            lock (Gate)
            {
                File.AppendAllText(AppPaths.LogFile, line);
            }
        }
        catch
        {
            // logging must never break merge/compare
        }
    }
}
