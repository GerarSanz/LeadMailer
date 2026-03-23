using System.IO;

namespace LeadMailer.Services;

/// <summary>Logger de archivo para diagnóstico en producción. Sin dependencias externas.</summary>
public static class AppLogger
{
    private static readonly string LogFolder =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                     "LeadMailer", "logs");

    public static void Info(string message)             => Write("INFO ",  message, null);
    public static void Warn(string message)             => Write("WARN ",  message, null);
    public static void Error(string message)            => Write("ERROR",  message, null);
    public static void Error(string message, Exception ex) => Write("ERROR", message, ex);

    private static void Write(string level, string message, Exception? ex)
    {
        try
        {
            Directory.CreateDirectory(LogFolder);
            var file = Path.Combine(LogFolder, $"LeadMailer_{DateTime.Now:yyyy-MM-dd}.log");
            var sb = new System.Text.StringBuilder();
            sb.Append($"{DateTime.Now:HH:mm:ss.fff} [{level}] {message}");
            if (ex != null)
            {
                sb.AppendLine();
                sb.Append($"  {ex.GetType().Name}: {ex.Message}");
                if (ex.StackTrace is { } st)
                {
                    sb.AppendLine();
                    sb.Append("  " + st.Replace(Environment.NewLine, Environment.NewLine + "  "));
                }
                if (ex.InnerException is { } inner)
                {
                    sb.AppendLine();
                    sb.Append($"  Inner: {inner.GetType().Name}: {inner.Message}");
                }
            }
            sb.AppendLine();
            File.AppendAllText(file, sb.ToString());
        }
        catch { /* No se puede loguear el error del propio logger */ }
    }
}
