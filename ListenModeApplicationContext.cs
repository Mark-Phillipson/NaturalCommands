using System;
using System.Drawing;
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

        public ListenModeApplicationContext()
        {
            _commands = new Commands(new HandleProcesses());

            var menu = new ContextMenuStrip();
            var openItem = new ToolStripMenuItem($"Open Voice Dictation ({HotkeyText})");
            openItem.Click += (_, __) => OpenVoiceDictation();

            var stopAutoClickItem = new ToolStripMenuItem("Stop Auto-Click");
            stopAutoClickItem.Click += (_, __) => StopAutoClick();

            var settingsItem = new ToolStripMenuItem("Settings...");
            settingsItem.Click += (_, __) => OpenSettingsForm();

            var exitItem = new ToolStripMenuItem("Exit");
            exitItem.Click += (_, __) => ExitThread();

            menu.Items.Add(openItem);
            menu.Items.Add(stopAutoClickItem);
            menu.Items.Add(settingsItem);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add(exitItem);

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
        }

        private void OpenVoiceDictation()
        {
            if (_dictationOpen) return;
            _dictationOpen = true;

            try
            {
                var text = Helpers.VoiceDictationHelper.ShowVoiceDictation(0);
                if (!string.IsNullOrWhiteSpace(text))
                {
                    text = Helpers.WordReplacementHelper.ApplyWordReplacements(text);
                    _commands.HandleNaturalAsync(text);
                }
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

        protected override void ExitThreadCore()
        {
            try { _hotkeyRegistrar.Dispose(); } catch { }
            try { TrayNotificationHelper.Dispose(); } catch { }
            
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
