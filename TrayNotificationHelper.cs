using System;
using System.Drawing;
using System.Windows.Forms;

namespace NaturalCommands
{
    public static class TrayNotificationHelper
    {
        private static NotifyIcon? _notifyIcon;
        private static Icon? _appIcon;
        private static ContextMenuStrip? _contextMenu;

        private static void EnsureNotifyIcon()
        {
            if (_notifyIcon != null) return;

            _notifyIcon = new NotifyIcon();
            _appIcon = SystemIcons.Information;
            _notifyIcon.Icon = _appIcon;
            _notifyIcon.Visible = true;
            _notifyIcon.Text = "NaturalCommands.NET";
        }

        public static void InitializeResidentTray(string tooltip, ContextMenuStrip menu, Icon? icon = null)
        {
            EnsureNotifyIcon();
            try
            {
                _contextMenu = menu;
                _notifyIcon!.ContextMenuStrip = _contextMenu;
            }
            catch { }

            try
            {
                if (icon != null)
                {
                    _appIcon = icon;
                    _notifyIcon!.Icon = _appIcon;
                }
            }
            catch { }

            SetTooltip(tooltip);
        }

        public static void SetTooltip(string tooltip)
        {
            EnsureNotifyIcon();
            try
            {
                // NotifyIcon.Text is limited (~63 chars); keep it short.
                if (tooltip.Length > 60) tooltip = tooltip.Substring(0, 60);
                var ni = _notifyIcon;
                if (ni == null) return;
                ni.Text = tooltip;
            }
            catch { }
        }

        public static void ShowNotification(string title, string message, int timeout = 5000)
        {
            EnsureNotifyIcon();
            var ni = _notifyIcon;
            if (ni == null) return;
            ni.BalloonTipTitle = title;
            ni.BalloonTipText = message;
            ni.BalloonTipIcon = ToolTipIcon.Info;
            ni.ShowBalloonTip(timeout);
        }

        public static void Dispose()
        {
            if (_notifyIcon != null)
            {
                _notifyIcon.Visible = false;
                _notifyIcon.Dispose();
                _notifyIcon = null;
            }

			try { _contextMenu?.Dispose(); } catch { }
			_contextMenu = null;
        }
    }
}
