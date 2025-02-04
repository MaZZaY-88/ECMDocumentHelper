using Serilog;
using System;

namespace ECMDocumentHelper.Helpers
{
    public static class Logger
    {
        static Logger()
        {
            Log.Logger = new LoggerConfiguration()
                .WriteTo.Console()
                .WriteTo.File("logs\\log-" + ".txt", rollingInterval: RollingInterval.Day, outputTemplate: "[{Timestamp:yyyy-MM-dd HH:mm:ss}] [{Level}] {Message}{NewLine}{Exception}")
                .CreateLogger();
        }

        public static void LogInformation(string message)
        {
            Log.Information(message);
        }

        public static void LogWarning(string message)
        {
            Log.Warning(message);
        }

        public static void LogError(string message, Exception ex = null)
        {
            if (ex != null)
            {
                Log.Error(ex, message);
            }
            else
            {
                Log.Error(message);
            }
        }

        public static void LogDebug(string message)
        {
            Log.Debug(message);
        }

        public static void LogFatal(string message, Exception ex = null)
        {
            if (ex != null)
            {
                Log.Fatal(ex, message);
            }
            else
            {
                Log.Fatal(message);
            }
        }

        public static void CloseAndFlush()
        {
            Log.CloseAndFlush();
        }
    }
}
