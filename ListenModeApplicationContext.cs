using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace NaturalCommands
{
    public sealed class ListenModeApplicationContext : ApplicationContext
    {
        public const string HotkeyText = "Win+Ctrl+H";

        // Path used by external processes to send ticker payloads to a running listen-mode instance.
        private static string GetTickerPayloadFilePath()
        {
            var baseDir = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "NaturalCommands");
            System.IO.Directory.CreateDirectory(baseDir);
            return System.IO.Path.Combine(baseDir, ".ticker_payload");
        }

        private readonly HotkeyRegistrar _hotkeyRegistrar;
        private readonly Commands _commands;
        private bool _dictationOpen;
        private SettingsForm? _settingsForm;
        private System.Windows.Forms.Timer? _tickerCheckTimer;
        private WindowsNotificationListenerService? _notificationListener;

        public ListenModeApplicationContext()
        {
            _commands = new Commands(new HandleProcesses());

            var menu = new ContextMenuStrip();
            menu.ShowImageMargin = true;

            var openItem = new ToolStripMenuItem($"Open Voice Dictation ({HotkeyText})");
            openItem.Click += (_, __) => OpenVoiceDictation();

            var stopAutoClickItem = new ToolStripMenuItem("Stop Auto-Click");
            stopAutoClickItem.Click += (_, __) => StopAutoClick();

            var settingsItem = new ToolStripMenuItem("Settings...");
            settingsItem.Click += (_, __) => OpenSettingsForm();

            var exitItem = new ToolStripMenuItem("Exit");
            exitItem.Click += (_, __) => ExitThread();

            // Generate and attach small icons for clarity in the tray menu
            Image? openIcon = null;
            Image? stopIcon = null;
            Image? settingsIcon = null;
            Image? exitIcon = null;
            try
            {
                openIcon = MenuIconGenerator.CreateMicrophoneImage();
                stopIcon = MenuIconGenerator.CreateStopImage();
                settingsIcon = MenuIconGenerator.CreateSettingsImage();
                exitIcon = MenuIconGenerator.CreateExitImage();

                openItem.Image = openIcon;
                stopAutoClickItem.Image = stopIcon;
                settingsItem.Image = settingsIcon;
                exitItem.Image = exitIcon;
            }
            catch { }

            menu.Items.Add(openItem);
            menu.Items.Add(stopAutoClickItem);
            menu.Items.Add(settingsItem);

            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(exitItem);

            // Dispose the generated images when the menu is disposed
            menu.Disposed += (_, __) =>
            {
                try { openIcon?.Dispose(); } catch { }
                try { stopIcon?.Dispose(); } catch { }
                try { settingsIcon?.Dispose(); } catch { }
                try { exitIcon?.Dispose(); } catch { }
            };

            // Use custom application icon
            var appIcon = AppIconGenerator.CreateAppIcon();
            TrayNotificationHelper.InitializeResidentTray($"NaturalCommands ({HotkeyText})", menu, appIcon);

            _hotkeyRegistrar = new HotkeyRegistrar();
            _hotkeyRegistrar.Activated += (_, __) => OpenVoiceDictation();

            bool registered = _hotkeyRegistrar.TryRegister(HotkeyModifiers.Win | HotkeyModifiers.Control | HotkeyModifiers.NoRepeat, Keys.H);
            if (!registered)
            {
                try
                {
                    TrayNotificationHelper.ShowNotification("Hotkey unavailable", $"{HotkeyText} is already in use. Use the tray menu instead.", 4000);
                }
                catch { }
            }

            // Start a ticker payload watcher so external processes can send ticker messages to the resident listen-mode instance.
            try
            {
                _tickerCheckTimer = new System.Windows.Forms.Timer { Interval = 500 };
                _tickerCheckTimer.Tick += (s, e) => { CheckForTickerPayload(); };
                _tickerCheckTimer.Start();
            }
            catch (Exception ex)
            {
                Helpers.Logger.LogWarning($"ListenMode: Failed to start ticker payload watcher: {ex.Message}");
            }

            // Start resident notification listener and forward notifications to the ticker
            try
            {
                _notificationListener = new WindowsNotificationListenerService();
                _notificationListener.NotificationsReceived += (lines) =>
                {
                    try
                    {
                        Helpers.Logger.LogInfo($"ListenMode: NotificationsReceived ({lines.Count} lines)");
                        ShowTickerLines(lines);
                    }
                    catch (Exception ex)
                    {
                        Helpers.Logger.LogError($"ListenMode: NotificationsReceived handler error: {ex.Message}");
                    }
                };

                System.Threading.Tasks.Task.Run(() =>
                {
                    try
                    {
                        var allowed = _notificationListener.Start();
                        if (!allowed)
                        {
                            Helpers.Logger.LogWarning("Notification listener access denied. Enable NaturalCommands in Windows Settings > System > Notifications.");
                        }
                        else
                        {
                            Helpers.Logger.LogInfo("Resident notification listener started.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Helpers.Logger.LogError($"Failed to start resident notification listener: {ex.Message}");
                    }
                });
            }
            catch (Exception ex)
            {
                Helpers.Logger.LogError($"ListenMode: Failed to initialize notification listener: {ex.Message}");
            }

        }

        private void OpenVoiceDictation()
        {
            if (_dictationOpen) return;
            _dictationOpen = true;

            try
            {
                // Use the event-based approach - form stays open for multiple commands
                Helpers.VoiceDictationHelper.ShowVoiceDictation(0, (commandText) =>
                {
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(commandText))
                        {
                            var processedText = Helpers.WordReplacementHelper.ApplyWordReplacements(commandText);
                            _commands.HandleNaturalAsync(processedText);
                        }
                    }
                    catch (Exception ex)
                    {
                        try { TrayNotificationHelper.ShowNotification("Command failed", ex.Message, 3500); } catch { }
                    }
                });
            }
            catch (Exception ex)
            {
                try { TrayNotificationHelper.ShowNotification("Voice dictation failed", ex.Message, 3500); } catch { }
            }
            finally
            {
                _dictationOpen = false;
            }
        }

        private void StopAutoClick()
        {
            try
            {
                var result = Helpers.AutoClickManager.Stop();
                TrayNotificationHelper.ShowNotification("Auto-Click", result, 2000);
            }
            catch (Exception ex)
            {
                TrayNotificationHelper.ShowNotification("Error", $"Failed to stop auto-click: {ex.Message}", 3000);
            }
        }

        private void OpenSettingsForm()
        {
            // Use single instance pattern - only one settings form at a time
            if (_settingsForm == null || _settingsForm.IsDisposed)
            {
                _settingsForm = new SettingsForm();
                _settingsForm.ShowDialog();
            }
            else
            {
                _settingsForm.BringToFront();
                _settingsForm.Activate();
            }
        }

        // Quick Clicks support removed. (Related methods deleted.)

        // Check for ticker payloads dropped by external processes (e.g., personal-assistant)
        private void CheckForTickerPayload()
        {
            var payloadFile = GetTickerPayloadFilePath();
            try
            {
                if (System.IO.File.Exists(payloadFile))
                {
                    var lines = System.IO.File.ReadAllLines(payloadFile).Select(l => l.Trim()).Where(l => !string.IsNullOrEmpty(l)).ToList();
                    System.IO.File.Delete(payloadFile);

                    Helpers.Logger.LogInfo($"ListenMode: Received ticker payload with {lines.Count} lines (file: {payloadFile})");

                    if (lines.Count > 0)
                    {
                        ShowTickerLines(lines);
                    }
                }
            }
            catch (Exception ex)
            {
                try { Helpers.Logger.LogError($"CheckForTickerPayload failed: {ex.Message}"); if (System.IO.File.Exists(payloadFile)) System.IO.File.Delete(payloadFile); } catch { }
            }
        }

        // Show the ticker lines in a new STA thread (re-usable by notification listener and ticker payload watcher)
        private void ShowTickerLines(System.Collections.Generic.List<string> lines)
        {
            // Similar to WindowsNotificationListenerService.ShowTicker
            lock (this)
            {
                try { Helpers.Logger.LogInfo($"ListenMode: ShowTickerLines called ({lines.Count} lines)"); } catch { }
                // Create a new form on its own STA thread so it can be invoked later for additions
                var linesCopy = new System.Collections.Generic.List<string>(lines);
                NaturalCommands.TickerOverlayForm? created = null;
                var handleReady = new System.Threading.ManualResetEventSlim(false);

                var thread = new System.Threading.Thread(() =>
                {
                    System.Windows.Forms.Application.EnableVisualStyles();
                    System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
                    var form = new NaturalCommands.TickerOverlayForm(linesCopy, cycleSeconds: 5, maxCycles: 0, topPosition: false);
                    form.FormClosed += (s, e) => { /* nothing to do */ };
                    var _ = form.Handle; // ensure handle created
                    created = form;
                    handleReady.Set();
                    System.Windows.Forms.Application.Run(form);
                });
                thread.SetApartmentState(System.Threading.ApartmentState.STA);
                thread.IsBackground = true;
                thread.Start();

                // Wait briefly for the form to be ready
                handleReady.Wait(TimeSpan.FromSeconds(3));
            }
        }

        // ForegroundMonitor support for Quick Clicks removed.

        protected override void ExitThreadCore()
        {
            try { _hotkeyRegistrar.Dispose(); } catch { }
            try { TrayNotificationHelper.Dispose(); } catch { }

            // Stop and dispose ticker check timer
            try
            {
                if (_tickerCheckTimer != null)
                {
                    _tickerCheckTimer.Stop();
                    _tickerCheckTimer.Dispose();
                }
            }
            catch { }
            
            // Kill all NaturalCommands processes to ensure clean shutdown
            try
            {
                var currentProcess = System.Diagnostics.Process.GetCurrentProcess();
                var allProcesses = System.Diagnostics.Process.GetProcessesByName("NaturalCommands");
                foreach (var process in allProcesses)
                {
                    if (process.Id != currentProcess.Id)
                    {
                        try { process.Kill(); } catch { }
                    }
                }
            }
            catch { }

            // Stop resident notification listener if running
            try
            {
                if (_notificationListener != null)
                {
                    try { _notificationListener.Stop(); } catch { }
                }
            }
            catch { }
            
            base.ExitThreadCore();
        }
    }
}
