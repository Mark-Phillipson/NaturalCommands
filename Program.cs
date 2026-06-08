using NaturalCommands;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Threading;
using System.Windows.Forms;

namespace ExecuteCommands_NET
{
    internal static class Program
    {
        private static bool HasOtherNaturalCommandsProcess()
        {
            var currentProcessId = Process.GetCurrentProcess().Id;
            return Process.GetProcessesByName("NaturalCommands").Any(process => process.Id != currentProcessId);
        }

        

        [STAThread]
        static void Main(string[] args)
        {
            // Minimal, safe startup to keep build/tests happy. Real CLI behavior is preserved elsewhere.
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                var argsText = args == null
                    ? "<null>"
                    : string.Join(" | ", args.Select((a, i) => $"[{i}]='{a}'"));
                NaturalCommands.Helpers.Logger.LogInfo($"Program.Main startup. Args: {argsText}. LogPath: {NaturalCommands.Helpers.Logger.LogPath}");
            }
            catch { }

            // Handle natural command mode: execute a single command and display any overlays
            try
            {
                string NormalizeModeToken(string value)
                {
                    var t = (value ?? string.Empty).Trim();
                    if (t.StartsWith("/") || t.StartsWith("-"))
                    {
                        t = t.Substring(1).Trim();
                    }
                    return t;
                }

                string NormalizeCommandSegment(string value)
                {
                    var t = (value ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(t))
                    {
                        return string.Empty;
                    }

                    // Talon may prefix command tokens with '/' (e.g. /identify activity)
                    if (t.StartsWith("/") || t.StartsWith("-"))
                    {
                        t = t.Substring(1).TrimStart();
                    }

                    return t;
                }

                var knownModes = new[] { "natural", "ticker", "ticker-file", "ticker-test", "export-vs-commands" };
                var cleanedArgs = (args ?? Array.Empty<string>())
                    .Select(a => (a ?? string.Empty).Trim())
                    .Where(a => !string.IsNullOrWhiteSpace(a))
                    .ToArray();

                // No special handling for the legacy 'listen' token — it is removed.

                var firstArg = cleanedArgs.Length > 0 ? cleanedArgs[0] : string.Empty;
                var firstModeToken = NormalizeModeToken(firstArg);
                var explicitNaturalMode = !string.IsNullOrEmpty(firstModeToken) && string.Equals(firstModeToken, "natural", StringComparison.OrdinalIgnoreCase);
                var implicitNaturalMode = !string.IsNullOrEmpty(firstModeToken) && !knownModes.Contains(firstModeToken, StringComparer.OrdinalIgnoreCase);
                var naturalMode = explicitNaturalMode || implicitNaturalMode;
                if (naturalMode)
                {
                    try
                    {
                        // Initialize UI context for overlays
                        VisualCandidateOverlayForm.InitializeUIContext();
                        
                        // Talon may invoke as either:
                        // 1) NaturalCommands.exe natural identify telegram
                        // 2) NaturalCommands.exe identify telegram
                        var commandSegments = explicitNaturalMode
                            ? cleanedArgs.Skip(1).Select(NormalizeCommandSegment)
                            : cleanedArgs.Select(NormalizeCommandSegment);
                        var commandText = string.Join(" ", commandSegments.Where(s => !string.IsNullOrWhiteSpace(s))).Trim();

                        try
                        {
                            NaturalCommands.Helpers.Logger.LogInfo($"Program.Main natural command text: '{commandText}' (explicitNaturalMode={explicitNaturalMode}, implicitNaturalMode={implicitNaturalMode}, firstModeToken='{firstModeToken}')");
                        }
                        catch { }
                        
                        // Create interpreter and execute
                        var interpreter = new NaturalLanguageInterpreter();
                        var result = interpreter.HandleNaturalAsync(commandText);
                        
                        // Keep a message pump alive for 15 seconds so UI overlays (candidates, debug)
                        // can display and auto-close. The overlay forms have their own timeouts (9s).
                        var exitTimer = new System.Timers.Timer(15000);
                        exitTimer.Elapsed += (s, e) =>
                        {
                            exitTimer.Stop();
                            exitTimer.Dispose();
                            Application.Exit();
                        };
                        exitTimer.AutoReset = false;
                        exitTimer.Start();

                        // UI message pump to allow overlays to render
                        while (exitTimer.Enabled)
                        {
                            Application.DoEvents();
                            System.Threading.Thread.Sleep(50);
                        }
                    }
                    catch (Exception ex)
                    {
                        try { NaturalCommands.Helpers.Logger.LogError($"Natural command error: {ex.Message}"); } catch { }
                    }
                    // Note: do not return here — continue startup so the tray/menu
                    // can be shown by default (resident behavior on startup).
                }
            }
            catch { }

            // Start resident tray/menu by default when enabled in settings.
            try
            {
                var settings = NaturalCommands.Models.AppSettings.Instance;
                var showTray = settings?.Notifications?.ShowTrayOnStartup ?? true;
                var testMode = settings?.Behavior?.TestMode ?? false;

                if (showTray && !testMode)
                {
                    if (HasOtherNaturalCommandsProcess())
                    {
                        try { NaturalCommands.Helpers.Logger.LogInfo("Skipping resident tray startup because another NaturalCommands process is already running."); } catch { }
                        return;
                    }

                    try
                    {
                        Application.Run(new ListenModeApplicationContext());
                    }
                    catch (Exception ex)
                    {
                        try { NaturalCommands.Helpers.Logger.LogError($"Failed to start ListenModeApplicationContext: {ex.Message}"); } catch { }
                    }

                    return;
                }
            }
            catch { }

        }

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private static string[]? ShowDebugInputDialog() => null;
    }
}
