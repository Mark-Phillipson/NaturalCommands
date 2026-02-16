using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using NaturalCommands.Models;
using System.Windows.Forms;

namespace NaturalCommands
{
    public sealed class ListenModeApplicationContext : ApplicationContext
    {
        public const string HotkeyText = "Win+Ctrl+H";

        private readonly HotkeyRegistrar _hotkeyRegistrar;
        private readonly Commands _commands;
        private bool _dictationOpen;
        private SettingsForm? _settingsForm;
        private System.Windows.Forms.Timer? _manualHideCheckTimer;

        public ListenModeApplicationContext()
        {
            // Clean up any leftover command file from previous run to prevent auto-triggering on startup
            try
            {
                var commandFile = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ".quick_clicks_command");
                if (System.IO.File.Exists(commandFile))
                {
                    System.IO.File.Delete(commandFile);
                    Helpers.Logger.LogDebug("ListenMode: Deleted leftover .quick_clicks_command file on startup");
                }
            }
            catch (Exception ex)
            {
                Helpers.Logger.LogWarning($"ListenMode: Failed to delete leftover command file: {ex.Message}");
            }

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

            // Quick Clicks submenu (Show / Edit / Create at Mouse / Open JSON)
            var quickClicksMenu = new ToolStripMenuItem("Quick Clicks");
            var qcShow = new ToolStripMenuItem("Show Quick Clicks");
            qcShow.Click += (_, __) => ShowQuickClicksForForegroundWindow();
            var qcEdit = new ToolStripMenuItem("Edit Quick Clicks");
            qcEdit.Click += (_, __) => EditQuickClicksForForegroundWindow();
            var qcCreateAtMouse = new ToolStripMenuItem("Create Quick Click at Mouse");
            qcCreateAtMouse.Click += (_, __) => { QuickClickOverlayForm.BeginPlacementAtCursorForForegroundWindow(); };
            var qcOpenJson = new ToolStripMenuItem("Open quick_clicks.json");
            qcOpenJson.Click += (_, __) => {
                try { var path = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "quick_clicks.json"); System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true }); } catch (Exception ex) { TrayNotificationHelper.ShowNotification("Open quick_clicks.json failed", ex.Message, 3000); }
            };
            quickClicksMenu.DropDownItems.Add(qcShow);
            quickClicksMenu.DropDownItems.Add(qcEdit);
            quickClicksMenu.DropDownItems.Add(qcCreateAtMouse);
            quickClicksMenu.DropDownItems.Add(new ToolStripSeparator());
            quickClicksMenu.DropDownItems.Add(qcOpenJson);
            menu.Items.Add(quickClicksMenu);

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

            // Start foreground monitor for Quick Clicks (auto-show/hide overlays)
            try
            {
                if (AppSettings.Instance.QuickClicks.Enabled)
                {
                    Helpers.ForegroundMonitor.ForegroundChanged += ForegroundMonitor_ForegroundChanged;
                    Helpers.ForegroundMonitor.Start();
                    
                    // Start timer to check for voice commands (sent via .quick_clicks_command file)
                    _manualHideCheckTimer = new System.Windows.Forms.Timer { Interval = 200 }; // Check every 200ms
                    _manualHideCheckTimer.Tick += (s, e) =>
                    {
                        // Check for manual hide flag
                        if (QuickClickOverlayForm.IsManuallyHidden && QuickClickOverlayForm.IsVisible)
                        {
                            Helpers.Logger.LogDebug("ListenMode: Manual hide flag detected, hiding overlay");
                            QuickClickOverlayForm.HideOverlay();
                        }
                        
                        // Check for voice commands
                        CheckForQuickClicksCommand();
                    };
                    _manualHideCheckTimer.Start();
                }
            }
            catch { }

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

        // Show the overlay for the currently focused window (used by tray menu commands)
        private void ShowQuickClicksForForegroundWindow()
        {
            try
            {
                // Clear manual hide flag when explicitly showing
                QuickClickOverlayForm.ClearManualHideFlag();
                
                var hwnd = Helpers.Win32ApiHelper.GetForegroundWindow();
                var proc = CurrentApplicationHelper.GetCurrentProcessName();
                var title = CurrentApplicationHelper.GetCurrentWindowTitle();

                int? mw = null, mh = null;
                Helpers.Win32ApiHelper.TryGetMonitorResolutionForWindow(hwnd, out var w, out var h);
                if (w > 0 && h > 0) { mw = w; mh = h; }

                var profilesForMonitor = Helpers.QuickClickLoader.GetProfilesForApp(proc ?? string.Empty, title, mw, mh).ToList();

                Models.QuickClickProfile? profileToShow = null;
                if (profilesForMonitor.Count > 0)
                {
                    profileToShow = profilesForMonitor.First();
                }
                else
                {
                    // Show null profile (warning banner) if no matching profile exists
                    profileToShow = null;
                }
                
                QuickClickOverlayForm.ShowOverlay(profileToShow, hwnd, enterEditMode: false);
                TrayNotificationHelper.ShowNotification("Quick Clicks", profileToShow != null ? "Overlay shown" : "No profile for this window", 2000);
            }
            catch (Exception ex)
            {
                try { Helpers.Logger.LogError($"ShowQuickClicksForForegroundWindow failed: {ex.Message}"); } catch { }
            }
        }

        // Edit quick clicks for the currently focused window (used by tray menu "Edit Quick Clicks")
        private void EditQuickClicksForForegroundWindow()
        {
            try
            {
                var hwnd = Helpers.Win32ApiHelper.GetForegroundWindow();
                var proc = CurrentApplicationHelper.GetCurrentProcessName();
                var title = CurrentApplicationHelper.GetCurrentWindowTitle();

                int? mw = null, mh = null;
                Helpers.Win32ApiHelper.TryGetMonitorResolutionForWindow(hwnd, out var w, out var h);
                if (w > 0 && h > 0) { mw = w; mh = h; }

                var profilesForMonitor = Helpers.QuickClickLoader.GetProfilesForApp(proc ?? string.Empty, title, mw, mh).ToList();

                Models.QuickClickProfile? profileToShow = null;
                if (profilesForMonitor.Count > 0)
                {
                    profileToShow = profilesForMonitor.First();
                }
                else
                {
                    // Create a new empty profile for this app
                    profileToShow = new Models.QuickClickProfile 
                    { 
                        ProcessName = proc ?? string.Empty, 
                        WindowTitlePattern = title, 
                        MonitorWidth = mw, 
                        MonitorHeight = mh,
                        Regions = new System.Collections.Generic.List<Models.QuickClickRegion>()
                    };
                }
                
                QuickClickOverlayForm.ShowOverlay(profileToShow, hwnd, enterEditMode: true);
                TrayNotificationHelper.ShowNotification("Quick Clicks", "Edit mode active", 2000);
            }
            catch (Exception ex)
            {
                Helpers.Logger.LogError($"EditQuickClicksForForegroundWindow failed: {ex.Message}");
                TrayNotificationHelper.ShowNotification("Quick Clicks Error", ex.Message, 3000);
            }
        }

        // Check for voice commands sent via .quick_clicks_command file
        private void CheckForQuickClicksCommand()
        {
            var commandFile = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ".quick_clicks_command");
            try
            {
                if (System.IO.File.Exists(commandFile))
                {
                    var command = System.IO.File.ReadAllText(commandFile).Trim().ToLowerInvariant();
                    System.IO.File.Delete(commandFile); // Delete immediately to avoid re-processing
                    
                    Helpers.Logger.LogInfo($"ListenMode: Processing quick clicks command: {command}");
                    
                    if (command.Contains("edit") && command.Contains("quick"))
                    {
                        EditQuickClicksForForegroundWindow();
                    }
                    else if (command.Contains("show") && command.Contains("quick"))
                    {
                        ShowQuickClicksForForegroundWindow();
                    }
                    else if (command.Contains("hide") && command.Contains("quick"))
                    {
                        QuickClickOverlayForm.HideOverlay();
                    }
                    else if (command.Contains("create") && command.Contains("quick"))
                    {
                        QuickClickOverlayForm.BeginPlacementAtCursorForForegroundWindow();
                    }
                    else
                    {
                        // Might be a region click command - pass to natural language interpreter
                        _commands.HandleNaturalAsync(command);
                    }
                }
            }
            catch (Exception ex)
            {
                try 
                { 
                    Helpers.Logger.LogError($"CheckForQuickClicksCommand failed: {ex.Message}");
                    if (System.IO.File.Exists(commandFile))
                        System.IO.File.Delete(commandFile);
                }
                catch { }
            }
        }

        // Foreground monitor event handler — decides which overlay (if any) to show for the focused app/window.
        private void ForegroundMonitor_ForegroundChanged(IntPtr hwnd, string? processName, string? windowTitle, int? monitorWidth, int? monitorHeight)
        {
            try
            {
                // Don't interfere if overlay is already visible in edit mode
                if (QuickClickOverlayForm.IsVisible && QuickClickOverlayForm.IsInEditMode)
                {
                    Helpers.Logger.LogDebug($"ForegroundMonitor: Skipping auto-show because overlay is in edit mode");
                    return;
                }
                
                // Don't auto-show if user manually hid the overlay
                if (QuickClickOverlayForm.IsManuallyHidden)
                {
                    Helpers.Logger.LogDebug($"ForegroundMonitor: Skipping auto-show because overlay was manually hidden");
                    return;
                }

                if (!AppSettings.Instance.QuickClicks.Enabled || !AppSettings.Instance.QuickClicks.AutoShowOnFocus)
                {
                    // Don't auto-hide if AutoShowOnFocus is disabled - just skip auto-showing
                    // (allows manual "show quick clicks" commands to work)
                    Helpers.Logger.LogDebug($"ForegroundMonitor: Skipping - QuickClicks.Enabled={AppSettings.Instance.QuickClicks.Enabled}, AutoShowOnFocus={AppSettings.Instance.QuickClicks.AutoShowOnFocus}");
                    return;
                }

                if (string.IsNullOrWhiteSpace(processName))
                {
                    QuickClickOverlayForm.HideOverlay();
                    return;
                }

                var profilesForMonitor = Helpers.QuickClickLoader.GetProfilesForApp(processName, windowTitle, monitorWidth, monitorHeight).ToList();
                var allProfilesForApp = Helpers.QuickClickLoader.GetProfilesForApp(processName, windowTitle, null, null).ToList();

                if (profilesForMonitor.Count > 0)
                {
                    NaturalCommands.Models.QuickClickProfile? exact = null;
                    if (monitorWidth.HasValue && monitorHeight.HasValue)
                    {
                        exact = profilesForMonitor.FirstOrDefault(p => p.MonitorWidth.HasValue && p.MonitorHeight.HasValue && p.MonitorWidth.Value == monitorWidth.Value && p.MonitorHeight.Value == monitorHeight.Value);
                    }

                    NaturalCommands.Models.QuickClickProfile profileToShow;
                    if (exact != null) profileToShow = exact;
                    else profileToShow = profilesForMonitor.First();
                    QuickClickOverlayForm.ShowOverlay(profileToShow, hwnd);
                    return;
                }

                if (allProfilesForApp.Count > 0)
                {
                    // There are profiles for this app but none that match the current monitor resolution -> show warning banner (null profile)
                    QuickClickOverlayForm.ShowOverlay(null, hwnd);
                    return;
                }

                QuickClickOverlayForm.HideOverlay();
            }
            catch (Exception ex)
            {
                try { Helpers.Logger.LogError($"ForegroundMonitor_ForegroundChanged: {ex.Message}"); } catch { }
            }
        }

        protected override void ExitThreadCore()
        {
            try { _hotkeyRegistrar.Dispose(); } catch { }
            try { TrayNotificationHelper.Dispose(); } catch { }

            // Ensure ForegroundMonitor is unsubscribed and stopped on shutdown
            try { Helpers.ForegroundMonitor.ForegroundChanged -= ForegroundMonitor_ForegroundChanged; } catch { }
            try { Helpers.ForegroundMonitor.Stop(); } catch { }
            
            // Stop and dispose manual hide check timer
            try 
            { 
                if (_manualHideCheckTimer != null)
                {
                    _manualHideCheckTimer.Stop();
                    _manualHideCheckTimer.Dispose();
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
            
            base.ExitThreadCore();
        }
    }
}
