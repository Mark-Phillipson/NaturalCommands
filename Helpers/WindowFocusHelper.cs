
using System;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace NaturalCommands.Helpers
{
    public static class WindowFocusHelper
    {
        // Helper methods to send key events
        public static void SendKeyDown(byte vk)
        {
            keybd_event(vk, 0, 0, 0);
        }

        public static void SendKeyUp(byte vk)
        {
            keybd_event(vk, 0, 2, 0);
        }
        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll")]
        static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);
        [StructLayout(LayoutKind.Sequential)]
        struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();
        private const string WindowsTerminalClass = "CASCADIA_HOSTING_WINDOW_CLASS";
        private const string ZoomWindowClass = "zoom_acc_notify_wnd"; // Actual Zoom main window class from user system
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("user32.dll")]
        static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
        [DllImport("user32.dll")]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int X, int Y);
        [DllImport("user32.dll")]
        static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, int dwExtraInfo);
        private const int SW_RESTORE = 9;
        const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        const uint MOUSEEVENTF_LEFTUP = 0x0004;

        public static bool FocusWindowByTitle(string titleSubstring)
        {
            IntPtr foundHwnd = IntPtr.Zero;
            string foundProcName = "";
            string foundClassName = "";
            string foundWindowTitle = "";

            // Helper to generate all singular/plural combinations for each word
            static IEnumerable<string> GetTitleVariants(string input)
            {
                var words = input.Split(' ');
                List<List<string>> wordVariants = new List<List<string>>();
                foreach (var word in words)
                {
                    var variants = new List<string> { word };
                    if (word.EndsWith("s") && word.Length > 1)
                    {
                        // Remove trailing 's' for singular
                        variants.Add(word.Substring(0, word.Length - 1));
                    }
                    else if (word.Length > 1)
                    {
                        // Add plural form
                        variants.Add(word + "s");
                    }
                    wordVariants.Add(variants);
                }
                // Generate all combinations
                var results = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                void Combine(int idx, List<string> current)
                {
                    if (idx == wordVariants.Count)
                    {
                        results.Add(string.Join(" ", current));
                        return;
                    }
                    foreach (var variant in wordVariants[idx])
                    {
                        current.Add(variant);
                        Combine(idx + 1, current);
                        current.RemoveAt(current.Count - 1);
                    }
                }
                Combine(0, new List<string>());
                return results;
            }

            var titleVariants = GetTitleVariants(titleSubstring);
            NaturalCommands.Helpers.Logger.LogDebug($"FocusWindowByTitle: Checking variants: {string.Join(", ", titleVariants)}");

            EnumWindows((hWnd, lParam) =>
            {
                if (!IsWindowVisible(hWnd)) return true;
                var title = new System.Text.StringBuilder(256);
                GetWindowText(hWnd, title, title.Capacity);
                string windowTitle = title.ToString();
                NaturalCommands.Helpers.Logger.LogDebug($"FocusWindowByTitle: Window title found: '{windowTitle}'");
                foreach (var variant in titleVariants)
                {
                    var variantWords = variant.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                    bool allWordsPresent = true;
                    foreach (var word in variantWords)
                    {
                        if (!windowTitle.Contains(word, StringComparison.OrdinalIgnoreCase))
                        {
                            allWordsPresent = false;
                            break;
                        }
                    }
                    if (allWordsPresent)
                    {
                        NaturalCommands.Helpers.Logger.LogDebug($"FocusWindowByTitle: Fuzzy matched variant '{variant}' in window title '{windowTitle}'");
                        foundHwnd = hWnd;
                        foundWindowTitle = windowTitle;
                        uint pid;
                        GetWindowThreadProcessId(hWnd, out pid);
                        try
                        {
                            var proc = Process.GetProcessById((int)pid);
                            foundProcName = proc.ProcessName;
                            var className = new System.Text.StringBuilder(256);
                            GetClassName(hWnd, className, className.Capacity);
                            foundClassName = className.ToString();
                        }
                        catch { }
                        return false; // Stop enumeration
                    }
                }
                return true;
            }, IntPtr.Zero);
            if (foundHwnd != IntPtr.Zero)
            {
                bool restoreResult = ShowWindow(foundHwnd, SW_RESTORE);
                if (!restoreResult)
                {
                    int err = Marshal.GetLastWin32Error();
                    NaturalCommands.Helpers.Logger.LogError($"ShowWindow failed. GetLastError={err}");
                }
                IntPtr foregroundHwnd = GetForegroundWindow();
                uint foregroundThread = GetWindowThreadProcessId(foregroundHwnd, out _);
                uint targetThread = GetWindowThreadProcessId(foundHwnd, out _);
                bool attachResult = AttachThreadInput(foregroundThread, targetThread, true);
                if (!attachResult)
                {
                    int err = Marshal.GetLastWin32Error();
                    NaturalCommands.Helpers.Logger.LogError($"AttachThreadInput failed. GetLastError={err}");
                }
                // Send dummy Alt key input to help bypass focus restrictions
                //keybd_event(0x12, 0, 0, 0); // Alt down
                //keybd_event(0x12, 0, 2, 0); // Alt up
                bool focusResult = SetForegroundWindow(foundHwnd);
                if (!focusResult)
                {
                    int err = Marshal.GetLastWin32Error();
                    NaturalCommands.Helpers.Logger.LogError($"SetForegroundWindow failed. GetLastError={err}");
                }
                // Check if foreground window actually changed
                IntPtr afterFocusHwnd = GetForegroundWindow();
                bool actuallyFocused = (afterFocusHwnd == foundHwnd);
                NaturalCommands.Helpers.Logger.LogDebug($"After SetForegroundWindow: foregroundHwnd={afterFocusHwnd}, expectedHwnd={foundHwnd}, actuallyFocused={actuallyFocused}");
                if (attachResult)
                    AttachThreadInput(foregroundThread, targetThread, false); // detach
                NaturalCommands.Helpers.Logger.LogDebug($"FocusWindowByTitle: hwnd={foundHwnd}, proc={foundProcName}, class={foundClassName}, title={foundWindowTitle}, ShowWindow(SW_RESTORE)={restoreResult}, AttachThreadInput={attachResult}, SetForegroundWindow={focusResult}");
                // Fallback: simulate mouse click if not focused (either SetForegroundWindow failed or window not actually foreground)
                if (!focusResult || !actuallyFocused)
                {
                    NaturalCommands.Helpers.Logger.LogDebug("FocusWindowByTitle: SetForegroundWindow failed or window not foreground, trying mouse click fallback.");
                    RECT rect;
                    if (GetWindowRect(foundHwnd, out rect))
                    {
                        int x = rect.Left + 10;
                        int y = rect.Top + 10;
                        SetCursorPos(x, y);
                        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, x, y, 0, 0);
                        NaturalCommands.Helpers.Logger.LogDebug($"FocusWindowByTitle: Simulated mouse click at ({x},{y}) on window.");
                        focusResult = SetForegroundWindow(foundHwnd);
                        if (!focusResult)
                        {
                            int err = Marshal.GetLastWin32Error();
                            NaturalCommands.Helpers.Logger.LogError($"SetForegroundWindow after mouse click failed. GetLastError={err}");
                        }
                        afterFocusHwnd = GetForegroundWindow();
                        actuallyFocused = (afterFocusHwnd == foundHwnd);
                        NaturalCommands.Helpers.Logger.LogDebug($"FocusWindowByTitle: SetForegroundWindow after mouse click: {focusResult}, actuallyFocused: {actuallyFocused}");
                    }
                }
                // Final fallback: show three focus apps (Ctrl+Alt+Tab) if still not focused
                if (!focusResult || !actuallyFocused)
                {
                    NaturalCommands.Helpers.Logger.LogDebug("FocusWindowByTitle: All focus attempts failed, sending Ctrl+Alt+Tab as last resort.");
                    // Send Ctrl+Alt+Tab key sequence
                    keybd_event(0x11, 0, 0, 0); // Ctrl down
                    keybd_event(0x12, 0, 0, 0); // Alt down
                    keybd_event(0x09, 0, 0, 0); // Tab down
                    keybd_event(0x09, 0, 2, 0); // Tab up
                    keybd_event(0x12, 0, 2, 0); // Alt up
                    keybd_event(0x11, 0, 2, 0); // Ctrl up
                    NaturalCommands.Helpers.Logger.LogDebug("FocusWindowByTitle: Sent Ctrl+Alt+Tab.");
                }
                return actuallyFocused;
            }
            else
            {
                NaturalCommands.Helpers.Logger.LogDebug($"FocusWindowByTitle: No matching window found. Last proc={foundProcName}, class={foundClassName}");
            }
            return false;
        }

        public static bool FocusZoom()
        {
            // For backward compatibility, call the generic method
            return FocusWindowByTitle("zoom");
        }

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("user32.dll")]
        private static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);
        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        public static bool FocusWindowsTerminal()
        {
            IntPtr foundHwnd = IntPtr.Zero;
            EnumWindows((hWnd, lParam) =>
            {
                if (!IsWindowVisible(hWnd)) return true;
                uint pid;
                GetWindowThreadProcessId(hWnd, out pid);
                try
                {
                    var proc = Process.GetProcessById((int)pid);
                    if (proc.ProcessName.Equals("wt", StringComparison.OrdinalIgnoreCase))
                    {
                        var className = new System.Text.StringBuilder(256);
                        GetClassName(hWnd, className, className.Capacity);
                        if (className.ToString().Equals(WindowsTerminalClass, StringComparison.OrdinalIgnoreCase))
                        {
                            foundHwnd = hWnd;
                            return false; // Stop enumeration
                        }
                    }
                }
                catch { }
                return true;
            }, IntPtr.Zero);
            if (foundHwnd != IntPtr.Zero)
            {
                return SetForegroundWindow(foundHwnd);
            }
            return false;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr FindWindow(string lpClassName, string? lpWindowName);

        /// <summary>
        /// Attempts to focus the Windows Taskbar. First tries to find the Shell_TrayWnd and set it foreground,
        /// falling back to sending Win+T if necessary.
        /// </summary>
        // Additional PInvoke for child enumeration
        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumChildProc lpEnumFunc, IntPtr lParam);
        private delegate bool EnumChildProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string? lpszClass, string? lpszWindow);

        public static bool FocusTaskbar()
        {
            try
            {
                IntPtr trayHwnd = FindWindow("Shell_TrayWnd", null);
                if (trayHwnd != IntPtr.Zero)
                {
                    try { ShowWindow(trayHwnd, SW_RESTORE); } catch { }

                    // First try AttachThreadInput + SetForegroundWindow like other focus helpers
                    IntPtr foregroundHwnd = GetForegroundWindow();
                    uint foregroundThread = GetWindowThreadProcessId(foregroundHwnd, out _);
                    uint trayThread = GetWindowThreadProcessId(trayHwnd, out _);
                    bool attachResult = false;
                    try
                    {
                        attachResult = AttachThreadInput(foregroundThread, trayThread, true);
                    }
                    catch { }

                    bool setResult = SetForegroundWindow(trayHwnd);
                    NaturalCommands.Helpers.Logger.LogDebug($"FocusTaskbar: Attempted SetForegroundWindow on Shell_TrayWnd: {setResult}, AttachThreadInput: {attachResult}");

                    if (attachResult)
                        AttachThreadInput(foregroundThread, trayThread, false);

                    if (setResult)
                        return true;

                    // As a more reliable fallback, try to find a toolbar child and simulate a click inside it to ensure taskbar becomes focused
                    IntPtr toolbarHwnd = IntPtr.Zero;
                    EnumChildWindows(trayHwnd, (h, l) =>
                    {
                        var className = new System.Text.StringBuilder(256);
                        GetClassName(h, className, className.Capacity);
                        var cn = className.ToString();
                        if (cn.IndexOf("MSTask", StringComparison.OrdinalIgnoreCase) >= 0 || cn.IndexOf("ToolbarWindow32", StringComparison.OrdinalIgnoreCase) >= 0 || cn.IndexOf("ReBarWindow32", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            toolbarHwnd = h;
                            return false; // stop enumeration
                        }
                        return true;
                    }, IntPtr.Zero);

                    if (toolbarHwnd == IntPtr.Zero)
                    {
                        // Try common FindWindowEx names
                        IntPtr maybe = FindWindowEx(trayHwnd, IntPtr.Zero, "MSTaskSwWClass", null);
                        if (maybe != IntPtr.Zero) toolbarHwnd = maybe;
                        maybe = FindWindowEx(trayHwnd, IntPtr.Zero, "TaskListThumbnailWnd", null);
                        if (maybe != IntPtr.Zero) toolbarHwnd = maybe;
                    }

                    if (toolbarHwnd != IntPtr.Zero)
                    {
                        RECT rect;
                        GetWindowRect(toolbarHwnd, out rect);
                        int cx = (rect.Left + rect.Right) / 2;
                        int cy = (rect.Top + rect.Bottom) / 2;

                        NaturalCommands.Helpers.Logger.LogDebug($"FocusTaskbar: Simulating click at ({cx},{cy}) within toolbar hwnd {toolbarHwnd}");

                        // simulate mouse click at center of toolbar
                        SetCursorPos(cx, cy);
                        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, cx, cy, 0, 0);

                        System.Threading.Thread.Sleep(100);

                        bool finalSet = SetForegroundWindow(trayHwnd);
                        NaturalCommands.Helpers.Logger.LogDebug($"FocusTaskbar: After simulated click SetForegroundWindow result: {finalSet}");
                        return finalSet;
                    }
                }

                // Fallback: send Win+T to focus the taskbar
                keybd_event(0x5B, 0, 0, 0); // Win down
                keybd_event(0x54, 0, 0, 0); // T down
                keybd_event(0x54, 0, 2, 0); // T up
                keybd_event(0x5B, 0, 2, 0); // Win up

                System.Threading.Thread.Sleep(120);

                IntPtr fg = GetForegroundWindow();
                if (fg != IntPtr.Zero)
                {
                    var classNameSb = new System.Text.StringBuilder(256);
                    GetClassName(fg, classNameSb, classNameSb.Capacity);
                    if (classNameSb.ToString().Equals("Shell_TrayWnd", StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            catch (Exception ex)
            {
                NaturalCommands.Helpers.Logger.LogError($"FocusTaskbar error: {ex.Message}");
            }
            return false;
        }

        /// <summary>
        /// Returns the HWND for the Windows Taskbar (Shell_TrayWnd) or IntPtr.Zero if not found.
        /// </summary>
        public static IntPtr GetTaskbarHwnd()
        {
            try
            {
                return FindWindow("Shell_TrayWnd", null);
            }
            catch { return IntPtr.Zero; }
        }

        /// <summary>
        /// Attempts to find the desktop SHELLDLL_DefView hwnd (which hosts desktop icons). If not found, returns the Progman hwnd or IntPtr.Zero.
        /// </summary>
        public static IntPtr GetDesktopHwnd()
        {
            try
            {
                // Try Progman first
                IntPtr progman = FindWindow("Progman", null);
                if (progman != IntPtr.Zero)
                {
                    IntPtr defView = FindWindowEx(progman, IntPtr.Zero, "SHELLDLL_DefView", null);
                    if (defView != IntPtr.Zero)
                        return defView;
                }

                // Search WorkerW windows for a SHELLDLL_DefView child
                IntPtr foundDefView = IntPtr.Zero;
                EnumWindows((hWnd, lParam) =>
                {
                    var className = new System.Text.StringBuilder(256);
                    GetClassName(hWnd, className, className.Capacity);
                    if (className.ToString().Equals("WorkerW", StringComparison.OrdinalIgnoreCase))
                    {
                        IntPtr child = FindWindowEx(hWnd, IntPtr.Zero, "SHELLDLL_DefView", null);
                        if (child != IntPtr.Zero)
                        {
                            foundDefView = child;
                            return false; // stop enumeration
                        }
                    }
                    return true;
                }, IntPtr.Zero);

                if (foundDefView != IntPtr.Zero)
                    return foundDefView;

                // Fallback to Progman itself
                if (progman != IntPtr.Zero)
                    return progman;

                return IntPtr.Zero;
            }
            catch { return IntPtr.Zero; }
        }

        /// <summary>
        /// Attempts to focus the desktop area by first sending Win+D (which reliably shows desktop icons), then locating
        /// the desktop list view (SysListView32) and simulating a click inside it to ensure focus. Falls back to additional
        /// Win+D retries if necessary.
        /// </summary>
        public static bool FocusDesktop()
        {
            try
            {
                // Send Win+D to show desktop (this generally works reliably)
                keybd_event(0x5B, 0, 0, 0); // Win down
                keybd_event(0x44, 0, 0, 0); // D down
                keybd_event(0x44, 0, 2, 0); // D up
                keybd_event(0x5B, 0, 2, 0); // Win up
                NaturalCommands.Helpers.Logger.LogDebug("FocusDesktop: Sent Win+D to show desktop.");
                System.Threading.Thread.Sleep(150);

                IntPtr desktopHwnd = GetDesktopHwnd();
                IntPtr listView = IntPtr.Zero;

                if (desktopHwnd != IntPtr.Zero)
                {
                    // Try to find SHELLDLL_DefView then SysListView32
                    IntPtr defView = FindWindowEx(desktopHwnd, IntPtr.Zero, "SHELLDLL_DefView", null);
                    if (defView != IntPtr.Zero)
                        listView = FindWindowEx(defView, IntPtr.Zero, "SysListView32", null);
                    else
                        listView = FindWindowEx(desktopHwnd, IntPtr.Zero, "SysListView32", null);
                }

                if (listView != IntPtr.Zero)
                {
                    // Click the center of the list view to ensure desktop receives focus
                    RECT rect;
                    if (GetWindowRect(listView, out rect))
                    {
                        int cx = (rect.Left + rect.Right) / 2;
                        int cy = (rect.Top + rect.Bottom) / 2;
                        NaturalCommands.Helpers.Logger.LogDebug($"FocusDesktop: Simulating click at ({cx},{cy}) inside SysListView32 hwnd {listView}");

                        SetCursorPos(cx, cy);
                        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, cx, cy, 0, 0);

                        System.Threading.Thread.Sleep(120);

                        IntPtr fg = GetForegroundWindow();
                        if (fg != IntPtr.Zero)
                        {
                            var classNameSb = new System.Text.StringBuilder(256);
                            GetClassName(fg, classNameSb, classNameSb.Capacity);
                            var cn = classNameSb.ToString();
                            NaturalCommands.Helpers.Logger.LogDebug($"FocusDesktop: After click foreground class={cn}");
                            if (cn.Equals("Progman", StringComparison.OrdinalIgnoreCase) || cn.Equals("WorkerW", StringComparison.OrdinalIgnoreCase) || cn.Equals("SHELLDLL_DefView", StringComparison.OrdinalIgnoreCase) || cn.Equals("SysListView32", StringComparison.OrdinalIgnoreCase))
                                return true;
                        }

                        // Try to set foreground to desktop window directly
                        if (desktopHwnd != IntPtr.Zero)
                        {
                            bool finalSet = SetForegroundWindow(desktopHwnd);
                            System.Threading.Thread.Sleep(50);
                            IntPtr after = GetForegroundWindow();
                            if (after == desktopHwnd) return true;
                        }

                        // As a last resort, try setting foreground on the listView itself
                        var parentSet = SetForegroundWindow(listView);
                        return parentSet;
                    }
                }

                // Final fallback: send Win+D again and consider it successful if foreground becomes Progman/WorkerW
                keybd_event(0x5B, 0, 0, 0); // Win down
                keybd_event(0x44, 0, 0, 0); // D down
                keybd_event(0x44, 0, 2, 0); // D up
                keybd_event(0x5B, 0, 2, 0); // Win up
                System.Threading.Thread.Sleep(120);
                IntPtr fg2 = GetForegroundWindow();
                if (fg2 != IntPtr.Zero)
                {
                    var classNameSb = new System.Text.StringBuilder(256);
                    GetClassName(fg2, classNameSb, classNameSb.Capacity);
                    var cn2 = classNameSb.ToString();
                    if (cn2.Equals("Progman", StringComparison.OrdinalIgnoreCase) || cn2.Equals("WorkerW", StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            catch (Exception ex)
            {
                NaturalCommands.Helpers.Logger.LogError($"FocusDesktop error: {ex.Message}");
            }
            return false;
        }
    }
}
