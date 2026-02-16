using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace NaturalCommands
{
    public static class CurrentApplicationHelper
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        public static string? GetCurrentProcessName()
        {
            IntPtr hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero)
                return null;
            int pid;
            GetWindowThreadProcessId(hwnd, out pid);
            try
            {
                Process proc = Process.GetProcessById(pid);
                return proc.ProcessName.ToLower();
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Returns the title of the current foreground window, or null when none.
        /// </summary>
        public static string? GetCurrentWindowTitle()
        {
            IntPtr hwnd = GetForegroundWindow();
            if (hwnd == IntPtr.Zero) return null;
            var sb = new System.Text.StringBuilder(1024);
            try
            {
                // Use centralized Win32 helper when available
                NaturalCommands.Helpers.Win32ApiHelper.GetWindowText(hwnd, sb, sb.Capacity);
                var title = sb.ToString();
                return string.IsNullOrWhiteSpace(title) ? null : title;
            }
            catch
            {
                return null;
            }
        }
    }
}
