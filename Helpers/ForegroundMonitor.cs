using System;
using System.Windows.Forms;

namespace NaturalCommands.Helpers
{
    /// <summary>
    /// Watches foreground-window changes (EVENT_SYSTEM_FOREGROUND) and raises a debounced
    /// ForegroundChanged event with process name, window title and monitor resolution.
    /// </summary>
    public static class ForegroundMonitor
    {
        private static IntPtr _hook = IntPtr.Zero;
        private static Win32ApiHelper.WinEventDelegate? _proc;
        private static IntPtr _latestHwnd = IntPtr.Zero;

        // Debounce rapid focus changes (150ms)
        private static readonly DebounceGate _debounce = new DebounceGate(150);
        private static System.Windows.Forms.Timer? _debounceTimer;

        /// <summary>
        /// Raised after a debounced foreground change. Parameters: hwnd, processName, windowTitle, monitorWidth, monitorHeight
        /// monitorWidth/monitorHeight may be null when unknown.
        /// </summary>
        public static event Action<IntPtr, string?, string?, int?, int?>? ForegroundChanged;

        public static void Start()
        {
            if (_hook != IntPtr.Zero) return;

            _proc = new Win32ApiHelper.WinEventDelegate(WinEventProc);
            _hook = Win32ApiHelper.SetWinEventHook(Win32ApiHelper.EVENT_SYSTEM_FOREGROUND, Win32ApiHelper.EVENT_SYSTEM_FOREGROUND, IntPtr.Zero, _proc, 0, 0, Win32ApiHelper.WINEVENT_OUTOFCONTEXT);

            _debounceTimer = new System.Windows.Forms.Timer { Interval = 100 };
            _debounceTimer.Tick += DebounceTimer_Tick;
            _debounceTimer.Start();
        }

        public static void Stop()
        {
            try
            {
                if (_hook != IntPtr.Zero)
                {
                    Win32ApiHelper.UnhookWinEvent(_hook);
                    _hook = IntPtr.Zero;
                }
            }
            catch { }

            try { _debounceTimer?.Stop(); } catch { }
            _debounceTimer = null;
            _proc = null;
        }

        private static void WinEventProc(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            if (eventType != Win32ApiHelper.EVENT_SYSTEM_FOREGROUND) return;
            _latestHwnd = hwnd;
            _debounce.MarkChange(Environment.TickCount64);
            if (_debounceTimer != null && !_debounceTimer.Enabled) _debounceTimer.Start();
        }

        private static void DebounceTimer_Tick(object? sender, EventArgs e)
        {
            if (!_debounce.TryConsume(Environment.TickCount64)) return;
            if (_debounceTimer != null) _debounceTimer.Stop();

            var hwnd = _latestHwnd;
            if (hwnd == IntPtr.Zero)
            {
                ForegroundChanged?.Invoke(IntPtr.Zero, null, null, null, null);
                return;
            }

            string? processName = CurrentApplicationHelper.GetCurrentProcessName();
            string? windowTitle = CurrentApplicationHelper.GetCurrentWindowTitle();

            int? monitorWidth = null;
            int? monitorHeight = null;
            try
            {
                IntPtr monitor = Win32ApiHelper.MonitorFromWindow(hwnd, 2 /*MONITOR_DEFAULTTONEAREST*/);
                if (monitor != IntPtr.Zero)
                {
                    var info = new Win32ApiHelper.MONITORINFOEX();
                    info.cbSize = System.Runtime.InteropServices.Marshal.SizeOf(typeof(Win32ApiHelper.MONITORINFOEX));
                    if (Win32ApiHelper.GetMonitorInfo(monitor, ref info))
                    {
                        monitorWidth = info.rcMonitor.Right - info.rcMonitor.Left;
                        monitorHeight = info.rcMonitor.Bottom - info.rcMonitor.Top;
                    }
                }
            }
            catch { /* ignore monitor resolution errors */ }

            ForegroundChanged?.Invoke(hwnd, processName, windowTitle, monitorWidth, monitorHeight);
        }
    }
}
