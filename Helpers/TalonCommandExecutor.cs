using System;
using System.Diagnostics;
using System.IO;

namespace NaturalCommands.Helpers
{
    public static class TalonCommandExecutor
    {
        public static string Execute(NaturalCommands.RunTalonCommandAction action)
        {
            var settings = Models.AppSettings.Instance.Talon;
            var dispatchCommand = SanitizeForDispatch(action.TalonCommand);

            if (!string.IsNullOrWhiteSpace(settings.CommandQueueFilePath))
            {
                try
                {
                    var queuePath = settings.CommandQueueFilePath.Trim();
                    var queueDir = Path.GetDirectoryName(queuePath);
                    if (!string.IsNullOrWhiteSpace(queueDir) && !Directory.Exists(queueDir))
                    {
                        Directory.CreateDirectory(queueDir);
                    }

                    File.AppendAllText(queuePath, dispatchCommand + Environment.NewLine);
                    Logger.LogInfo($"TalonCommandExecutor: queued '{dispatchCommand}' to '{queuePath}'.");
                    return $"Queued Talon command: {dispatchCommand}";
                }
                catch (Exception ex)
                {
                    Logger.LogError($"TalonCommandExecutor: failed writing queue file '{settings.CommandQueueFilePath}'. {ex.Message}");
                }
            }

            if (!string.IsNullOrWhiteSpace(settings.BridgeExecutable))
            {
                try
                {
                    var argsTemplate = string.IsNullOrWhiteSpace(settings.BridgeArgumentsTemplate)
                        ? "{command_quoted}"
                        : settings.BridgeArgumentsTemplate;

                    var args = argsTemplate
                        .Replace("{command}", dispatchCommand, StringComparison.Ordinal)
                        .Replace("{command_quoted}", QuoteArgument(dispatchCommand), StringComparison.Ordinal);

                    var psi = new ProcessStartInfo(settings.BridgeExecutable, args)
                    {
                        UseShellExecute = true
                    };

                    Process.Start(psi);
                    Logger.LogInfo($"TalonCommandExecutor: dispatched '{dispatchCommand}' via bridge '{settings.BridgeExecutable}'.");
                    return $"Dispatched Talon command: {dispatchCommand}";
                }
                catch (Exception ex)
                {
                    Logger.LogError($"TalonCommandExecutor: bridge execution failed. {ex.Message}");
                    return $"Matched Talon command but dispatch failed: {ex.Message}";
                }
            }

            return "Matched Talon command but no dispatch bridge is configured. Set Talon.CommandQueueFilePath or Talon.BridgeExecutable in settings.json.";
        }

        private static string QuoteArgument(string value)
        {
            if (string.IsNullOrEmpty(value)) return "\"\"";
            return "\"" + value.Replace("\"", "\\\"") + "\"";
        }

        private static string SanitizeForDispatch(string command)
        {
            if (string.IsNullOrWhiteSpace(command))
            {
                return string.Empty;
            }

            var sanitized = command.Trim();

            while (sanitized.StartsWith("^", StringComparison.Ordinal))
            {
                sanitized = sanitized.Substring(1).TrimStart();
            }

            while (sanitized.EndsWith("$", StringComparison.Ordinal))
            {
                sanitized = sanitized.Substring(0, sanitized.Length - 1).TrimEnd();
            }

            sanitized = sanitized.Replace("[", " ").Replace("]", " ");
            sanitized = string.Join(" ", sanitized.Split(' ', StringSplitOptions.RemoveEmptyEntries));

            return string.IsNullOrWhiteSpace(sanitized) ? command.Trim() : sanitized;
        }
    }
}
