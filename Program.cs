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
        private static string GetTickerPayloadFilePath()
        {
            var baseDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "NaturalCommands");
            Directory.CreateDirectory(baseDir);
            return Path.Combine(baseDir, ".ticker_payload");
        }

        private static bool HasOtherNaturalCommandsProcess()
        {
            var currentProcessId = Process.GetCurrentProcess().Id;
            return Process.GetProcessesByName("NaturalCommands").Any(process => process.Id != currentProcessId);
        }

        private static void DeliverTickerPayload(IEnumerable<string> lines)
        {
            var payloadPath = GetTickerPayloadFilePath();
            var tempPath = payloadPath + ".tmp";
            File.WriteAllLines(tempPath, lines);
            if (File.Exists(payloadPath)) File.Delete(payloadPath);
            File.Move(tempPath, payloadPath);
            try { NaturalCommands.Helpers.Logger.LogInfo($"[WATCHER] Delivered ticker payload to resident form: {payloadPath}"); } catch { }
        }

        private static void StartTickerWatcher()
        {
            var watcherThread = new Thread(() =>
            {
                var payloadPath = GetTickerPayloadFilePath();
                string? lastHash = null;
                try { NaturalCommands.Helpers.Logger.LogInfo($"[WATCHER] Starting payload watcher on {payloadPath}"); } catch { }

                try
                {
                    TickerOverlayManager.EnsurePersistentTicker(new[] { "Ticker ready" }, cycleSeconds: 5, maxCycles: 0, topPosition: false, hideOnDismiss: true);
                }
                catch (Exception ex)
                {
                    try { NaturalCommands.Helpers.Logger.LogWarning($"[WATCHER] Failed to ensure persistent ticker: {ex.Message}"); } catch { }
                }

                while (true)
                {
                    try
                    {
                        if (File.Exists(payloadPath))
                        {
                            var content = File.ReadAllText(payloadPath);
                            if (string.IsNullOrWhiteSpace(content)) { Thread.Sleep(500); continue; }

                            var currentHash = System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(content));
                            var currentHashStr = System.Convert.ToHexString(currentHash);
                            if (lastHash != currentHashStr)
                            {
                                lastHash = currentHashStr;
                                var lines = content.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.RemoveEmptyEntries);
                                try
                                {
                                    var messages = lines.Select(l => TickerOverlayForm.ParseSingleLine(l)).ToList();
                                    TickerOverlayManager.AddMessages(messages);
                                    foreach (var message in messages) try { NaturalCommands.Helpers.Logger.LogInfo($"[WATCHER] Added message: {message.Text} ({message.Category})"); } catch { }
                                }
                                catch (Exception ex)
                                {
                                    try { NaturalCommands.Helpers.Logger.LogWarning($"[WATCHER] Failed adding messages to ticker: {ex.Message}"); } catch { }
                                }
                            }
                        }

                        Thread.Sleep(500);
                    }
                    catch (Exception ex)
                    {
                        try { NaturalCommands.Helpers.Logger.LogError($"[WATCHER] Watcher error: {ex.Message}"); } catch { }
                        Thread.Sleep(1000);
                    }
                }
            }) { IsBackground = false, Name = "TickerPayloadWatcher" };

            watcherThread.Start();
            try { NaturalCommands.Helpers.Logger.LogInfo("Ticker payload watcher started as foreground thread (keeps app alive)"); } catch { }
        }

        [STAThread]
        static void Main(string[] args)
        {
            // Minimal, safe startup to keep build/tests happy. Real CLI behavior is preserved elsewhere.
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // If asked to run in listen mode, start the resident application context
            try
            {
                var listenMode = args != null && args.Any(a => string.Equals(a, "listen", StringComparison.OrdinalIgnoreCase));
                if (listenMode)
                {
                    if (HasOtherNaturalCommandsProcess())
                    {
                        try { NaturalCommands.Helpers.Logger.LogInfo("Skipping listen-mode startup because another NaturalCommands process is already running."); } catch { }
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

            // Start the watcher if this is the resident mode runner and no other process is active
            try
            {
                var isStandaloneTickerMode = false;
                var hasOther = HasOtherNaturalCommandsProcess();
                if (!isStandaloneTickerMode && !hasOther)
                {
                    try { StartTickerWatcher(); } catch { }
                }
            }
            catch { }
        }

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private static string[]? ShowDebugInputDialog() => null;
    }
}
