using System;
using System.Runtime.InteropServices;

namespace NaturalCommands.Helpers
{
    // Provides public access to Win32 API functions and structs
    public static class Win32ApiHelper
    {
        public const int GWL_STYLE = -16;
        public const int WS_MAXIMIZEBOX = 0x00010000;

        [DllImport("user32.dll", SetLastError = true)]
        public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        // MonitorFromWindow flags (used when resolving which monitor contains a window)
        public const uint MONITOR_DEFAULTTONULL = 0x00000000;
        public const uint MONITOR_DEFAULTTOPRIMARY = 0x00000001;
        public const uint MONITOR_DEFAULTTONEAREST = 0x00000002;

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFOEX lpmi);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool GetWindowRect(IntPtr hWnd, ref RECT lpRect);

        // WinEvent hook support (used by ForegroundMonitor)
        public const uint EVENT_SYSTEM_FOREGROUND = 0x0003;
        public const uint WINEVENT_OUTOFCONTEXT = 0x0000;

        public delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        [DllImport("user32.dll")]
        public static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);

        [DllImport("user32.dll")]
        public static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        // Convenience wrapper: return window title (null when empty)
        public static string? GetWindowTitle(IntPtr hWnd)
        {
            var sb = new System.Text.StringBuilder(512);
            int len = GetWindowText(hWnd, sb, sb.Capacity);
            if (len > 0) return sb.ToString();
            return null;
        }

        // Try to resolve the monitor resolution that contains the given window handle.
        // Returns width/height in pixels (screen coordinates) when successful.
        public static bool TryGetMonitorResolutionForWindow(IntPtr hwnd, out int width, out int height)
        {
            width = height = 0;
            IntPtr hMonitor = MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST);
            if (hMonitor == IntPtr.Zero) return false;
            var info = new MONITORINFOEX();
            info.cbSize = Marshal.SizeOf(typeof(MONITORINFOEX));
            if (!GetMonitorInfo(hMonitor, ref info)) return false;
            width = info.rcMonitor.Right - info.rcMonitor.Left;
            height = info.rcMonitor.Bottom - info.rcMonitor.Top;
            return true;
        }

        // Try to get resolution directly from a monitor handle.
        public static bool TryGetMonitorResolution(IntPtr hMonitor, out int width, out int height)
        {
            width = height = 0;
            if (hMonitor == IntPtr.Zero) return false;
            var info = new MONITORINFOEX();
            info.cbSize = Marshal.SizeOf(typeof(MONITORINFOEX));
            if (!GetMonitorInfo(hMonitor, ref info)) return false;
            width = info.rcMonitor.Right - info.rcMonitor.Left;
            height = info.rcMonitor.Bottom - info.rcMonitor.Top;
            return true;
        }

        // Helper to enumerate all monitor handles
        public static IEnumerable<IntPtr> GetAllMonitors()
        {
            var monitors = new List<IntPtr>();
            bool MonitorEnum(IntPtr hMonitor, IntPtr hdcMonitor, ref RECT lprcMonitor, IntPtr dwData)
            {
                monitors.Add(hMonitor);
                return true;
            }
            EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero, MonitorEnum, IntPtr.Zero);
            return monitors;
        }

        [DllImport("user32.dll")]
        private static extern bool EnumDisplayMonitors(IntPtr hdc, IntPtr lprcClip, MonitorEnumProc lpfnEnum, IntPtr dwData);
        private delegate bool MonitorEnumProc(IntPtr hMonitor, IntPtr hdcMonitor, ref RECT lprcMonitor, IntPtr dwData);

        // Structs
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        public struct MONITORINFOEX
        {
            public int cbSize;
            public RECT rcMonitor;
            public RECT rcWork;
            public uint dwFlags;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
            public string szDevice;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
    }
}
