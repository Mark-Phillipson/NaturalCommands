using System;
using System.IO;

namespace NaturalCommands.Helpers
{
    // Simple centralized logger used by the app. Writes to a single file which can
    // be overridden by setting the NATURALCOMMANDS_LOG environment variable.
    public static class Logger
    {
        private static readonly string s_logPath;

        static Logger()
        {
            // Allow overriding the log file location via env var (useful for tests or alternate installs)
            var env = Environment.GetEnvironmentVariable("NATURALCOMMANDS_LOG");
            if (!string.IsNullOrWhiteSpace(env))
            {
                try { s_logPath = Path.GetFullPath(env); } catch { s_logPath = env; }
            }
            else
            {
                // Default to the repo layout used during development/publish: ../.. /bin/app.log
                var p = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "bin", "app.log");
                s_logPath = Path.GetFullPath(p);
            }
        }

        public static string LogPath => s_logPath;

        public static void EnsureLogDirExists()
        {
            try
            {
                var dir = Path.GetDirectoryName(s_logPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);
            }
            catch { }
        }

        public static void AppendLog(string message)
        {
            try
            {
                EnsureLogDirExists();
                using var fs = new FileStream(s_logPath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
                using var sw = new StreamWriter(fs);
                sw.Write(message);
            }
            catch { }
        }

        public static void AppendLine(string message)
            => AppendLog(message.EndsWith("\n") ? message : message + "\n");

        public static void AppendFormat(string fmt, params object[] args)
            => AppendLog(string.Format(fmt, args));

        public static void Log(string message)
        {
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            AppendLog($"[{timestamp}] {message}{Environment.NewLine}");
        }

        public static void LogInfo(string message)
        {
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            AppendLog($"[{timestamp}] [INFO] {message}{Environment.NewLine}");
        }

        public static void LogDebug(string message)
        {
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            AppendLog($"[{timestamp}] [DEBUG] {message}{Environment.NewLine}");
        }

        public static void LogWarning(string message)
        {
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            AppendLog($"[{timestamp}] [WARN] {message}{Environment.NewLine}");
        }

        public static void LogError(string message, Exception? ex = null)
        {
            string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            string errorMsg = $"[{timestamp}] [ERROR] {message}";
            if (ex != null)
            {
                errorMsg += $"{Environment.NewLine}{ex}";
            }
            AppendLog($"{errorMsg}{Environment.NewLine}");
        }
    }
}
