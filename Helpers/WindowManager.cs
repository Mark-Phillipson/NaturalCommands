using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
namespace NaturalCommands.Helpers
{
    // Handles window management actions (maximize, move, always on top, etc.)
    public class WindowManager
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [StructLayout(LayoutKind.Sequential)]
        private struct WINDOWPLACEMENT
        {
            public int length;
            public int flags;
            public int showCmd;
            public System.Drawing.Point ptMinPosition;
            public System.Drawing.Point ptMaxPosition;
            public System.Drawing.Rectangle rcNormalPosition;
        }

        private const int SW_MAXIMIZE = 3;
        private const int SW_SHOWNORMAL = 1;
        private const int WS_MAXIMIZE = 0x01000000;

        public static string ExecuteMoveWindow(MoveWindowAction move)
        {
            // Get active window handle
            IntPtr hWnd = NaturalCommands.Commands.GetForegroundWindow();
            
            // Log the window handle info for debugging
            var className = new System.Text.StringBuilder(256);
            Win32ApiHelper.GetClassName(hWnd, className, className.Capacity);
            Logger.LogDebug($"WindowManager: hWnd={hWnd}, ClassName={className}");
            
            // Maximize logic
            if ((move.Position == "center" || move.Position == null) && move.WidthPercent == 100 && move.HeightPercent == 100 && move.Monitor != "next")
            {
                // Check if this is our own process window (console) - if so, skip maximizing it
                uint processId = 0;
                GetWindowThreadProcessId(hWnd, out processId);
                uint currentProcessId = (uint)System.Diagnostics.Process.GetCurrentProcess().Id;
                Logger.LogDebug($"WindowManager: Window ProcessId={processId}, CurrentProcessId={currentProcessId}");
                
                if (processId == currentProcessId)
                {
                    Logger.LogWarning("WindowManager: Attempted to maximize own process window - this is likely incorrect");
                    return "Cannot maximize: the foreground window is the NaturalCommands process itself.";
                }
                
                Win32ApiHelper.SetForegroundWindow(hWnd);
                int style = Win32ApiHelper.GetWindowLong(hWnd, Win32ApiHelper.GWL_STYLE);
                bool canMaximize = (style & Win32ApiHelper.WS_MAXIMIZEBOX) != 0;
                Logger.LogDebug($"WindowManager: style={style:X8}, canMaximize={canMaximize}");
                
                if (!canMaximize)
                {
                    return "Window cannot be maximized (missing maximize button).";
                }
                
                // Get monitor working area - use Screen class as fallback since GetMonitorInfo can fail
                Win32ApiHelper.RECT windowRect = new Win32ApiHelper.RECT();
                Win32ApiHelper.GetWindowRect(hWnd, ref windowRect);
                
                System.Drawing.Rectangle workArea;
                try
                {
                    var screen = Screen.FromHandle(hWnd);
                    workArea = screen.WorkingArea;
                }
                catch
                {
                    // Fallback to primary screen
                    workArea = Screen.PrimaryScreen?.WorkingArea ?? new System.Drawing.Rectangle(0, 0, 1920, 1080);
                }
                
                Logger.LogDebug($"WindowManager: Window rect=({windowRect.Left},{windowRect.Top},{windowRect.Right},{windowRect.Bottom})");
                Logger.LogDebug($"WindowManager: Work area=({workArea.Left},{workArea.Top},{workArea.Right},{workArea.Bottom})");
                
                // A window is truly maximized if it covers (or exceeds) the working area
                // Allow a few pixels tolerance for window borders
                const int tolerance = 20;
                bool isVisuallyMaximized = 
                    windowRect.Left <= workArea.Left + tolerance &&
                    windowRect.Top <= workArea.Top + tolerance &&
                    windowRect.Right >= workArea.Right - tolerance &&
                    windowRect.Bottom >= workArea.Bottom - tolerance;
                
                Logger.LogDebug($"WindowManager: isVisuallyMaximized={isVisuallyMaximized}");
                
                if (isVisuallyMaximized)
                {
                    return "Window is already maximized.";
                }
                
                // Try ShowWindow first
                Win32ApiHelper.ShowWindow(hWnd, SW_MAXIMIZE);
                
                // Give the window a moment to process
                System.Threading.Thread.Sleep(50);
                
                // Verify the window was actually maximized by checking rect again
                Win32ApiHelper.RECT windowRectAfter = new Win32ApiHelper.RECT();
                Win32ApiHelper.GetWindowRect(hWnd, ref windowRectAfter);
                bool isMaximizedAfter = 
                    windowRectAfter.Left <= workArea.Left + tolerance &&
                    windowRectAfter.Top <= workArea.Top + tolerance &&
                    windowRectAfter.Right >= workArea.Right - tolerance &&
                    windowRectAfter.Bottom >= workArea.Bottom - tolerance;
                Logger.LogDebug($"WindowManager: After ShowWindow rect=({windowRectAfter.Left},{windowRectAfter.Top},{windowRectAfter.Right},{windowRectAfter.Bottom}), isMaximizedAfter={isMaximizedAfter}");
                
                if (isMaximizedAfter)
                {
                    return "Window maximized.";
                }
                
                // ShowWindow didn't work (common with modern apps like Windows Terminal)
                // Fall back to keyboard shortcut: Win+Up (may need to press twice for snap->maximize)
                Logger.LogDebug("WindowManager: ShowWindow failed, trying Win+Up keyboard shortcut");
                try
                {
                    var sim = new WindowsInput.InputSimulator();
                    
                    // Press Win+Up up to 2 times (first may snap to top, second maximizes)
                    for (int attempt = 0; attempt < 2; attempt++)
                    {
                        sim.Keyboard.ModifiedKeyStroke(
                            WindowsInput.Native.VirtualKeyCode.LWIN,
                            WindowsInput.Native.VirtualKeyCode.UP);
                        
                        // Give it a moment to process
                        System.Threading.Thread.Sleep(150);
                        
                        // Check if maximized
                        Win32ApiHelper.GetWindowRect(hWnd, ref windowRectAfter);
                        isMaximizedAfter = 
                            windowRectAfter.Left <= workArea.Left + tolerance &&
                            windowRectAfter.Top <= workArea.Top + tolerance &&
                            windowRectAfter.Right >= workArea.Right - tolerance &&
                            windowRectAfter.Bottom >= workArea.Bottom - tolerance;
                        Logger.LogDebug($"WindowManager: After Win+Up (attempt {attempt + 1}) rect=({windowRectAfter.Left},{windowRectAfter.Top},{windowRectAfter.Right},{windowRectAfter.Bottom}), isMaximizedAfter={isMaximizedAfter}");
                        
                        if (isMaximizedAfter)
                        {
                            return "Window maximized.";
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogError($"WindowManager: Win+Up keyboard shortcut failed: {ex.Message}");
                }
                
                return "Failed to maximize window.";
            }
            // Move window to left half
            if (move.Position == "left" && move.WidthPercent == 50 && move.HeightPercent == 100)
            {
                IntPtr monitor = Win32ApiHelper.MonitorFromWindow(hWnd, 2 /*MONITOR_DEFAULTTONEAREST*/);
                Win32ApiHelper.MONITORINFOEX info = new Win32ApiHelper.MONITORINFOEX();
                info.cbSize = System.Runtime.InteropServices.Marshal.SizeOf(typeof(Win32ApiHelper.MONITORINFOEX));
                bool gotInfo = monitor != IntPtr.Zero && Win32ApiHelper.GetMonitorInfo(monitor, ref info);
                if (!gotInfo)
                {
                    try
                    {
                        NaturalCommands.Helpers.Logger.LogWarning("WindowManager: failed to get monitor info, falling back to primary screen");
                    }
                    catch { }
                    // fallback to primary screen working area
                    try
                    {
                        var wa = Screen.PrimaryScreen?.WorkingArea ?? new System.Drawing.Rectangle(0, 0, SystemInformation.PrimaryMonitorSize.Width, SystemInformation.PrimaryMonitorSize.Height);
                        info.rcWork.Left = wa.Left;
                        info.rcWork.Top = wa.Top;
                        info.rcWork.Right = wa.Right;
                        info.rcWork.Bottom = wa.Bottom;
                        gotInfo = true;
                    }
                    catch
                    {
                        return "Failed to get monitor info.";
                    }
                }
                var rect = info.rcWork;
                int width = (rect.Right - rect.Left) / 2;
                int height = rect.Bottom - rect.Top;
                int x = rect.Left;
                int y = rect.Top;
                bool success = Win32ApiHelper.SetWindowPos(hWnd, IntPtr.Zero, x, y, width, height, 0x0040 /*SWP_SHOWWINDOW*/);
                if (!success)
                {
                    int error = System.Runtime.InteropServices.Marshal.GetLastWin32Error();
                    return $"Failed to move window left. Win32 error: {error}";
                }
                return "Window moved to left half.";
            }
            // Move window to right half
            if (move.Position == "right" && move.WidthPercent == 50 && move.HeightPercent == 100)
            {
                IntPtr monitor = Win32ApiHelper.MonitorFromWindow(hWnd, 2 /*MONITOR_DEFAULTTONEAREST*/);
                Win32ApiHelper.MONITORINFOEX info = new Win32ApiHelper.MONITORINFOEX();
                info.cbSize = System.Runtime.InteropServices.Marshal.SizeOf(typeof(Win32ApiHelper.MONITORINFOEX));
                bool gotInfo = monitor != IntPtr.Zero && Win32ApiHelper.GetMonitorInfo(monitor, ref info);
                if (!gotInfo)
                {
                    try
                    {
                        NaturalCommands.Helpers.Logger.LogWarning("WindowManager: failed to get monitor info for right half, falling back to primary screen");
                    }
                    catch { }
                    try
                    {
                        var wa = Screen.PrimaryScreen?.WorkingArea ?? new System.Drawing.Rectangle(0, 0, SystemInformation.PrimaryMonitorSize.Width, SystemInformation.PrimaryMonitorSize.Height);
                        info.rcWork.Left = wa.Left;
                        info.rcWork.Top = wa.Top;
                        info.rcWork.Right = wa.Right;
                        info.rcWork.Bottom = wa.Bottom;
                        gotInfo = true;
                    }
                    catch
                    {
                        return "Failed to get monitor info.";
                    }
                }
                var rect = info.rcWork;
                int width = (rect.Right - rect.Left) / 2;
                int height = rect.Bottom - rect.Top;
                int x = rect.Left + width;
                int y = rect.Top;
                bool success = Win32ApiHelper.SetWindowPos(hWnd, IntPtr.Zero, x, y, width, height, 0x0040 /*SWP_SHOWWINDOW*/);
                if (!success)
                {
                    int error = System.Runtime.InteropServices.Marshal.GetLastWin32Error();
                    return $"Failed to move window right. Win32 error: {error}";
                }
                return "Window moved to right half.";
            }
            // Move window to other monitor
            if (move.Monitor == "next" && (move.WidthPercent == 0 || move.WidthPercent == null) && (move.HeightPercent == 0 || move.HeightPercent == null))
            {
                IntPtr activeHWnd = NaturalCommands.Commands.GetForegroundWindow();
                if (activeHWnd == IntPtr.Zero)
                {
                    return "No active window found.";
                }
                Screen? currentScreen;
                try
                {
                    currentScreen = Screen.FromHandle(activeHWnd);
                }
                catch
                {
                    currentScreen = Screen.PrimaryScreen ?? (Screen.AllScreens.Length > 0 ? Screen.AllScreens[0] : null);
                }

                if (currentScreen == null)
                {
                    return "No screens detected.";
                }

                Screen[] allScreens = Screen.AllScreens;
                Screen nextScreen = currentScreen; // default to current if we can't find another
                for (int i = 0; i < allScreens.Length; i++)
                {
                    if (allScreens[i].DeviceName == currentScreen.DeviceName)
                    {
                        nextScreen = allScreens[(i + 1) % allScreens.Length];
                        break;
                    }
                }
                // Get current window size
                Win32ApiHelper.RECT currentRect = new Win32ApiHelper.RECT();
                Win32ApiHelper.GetWindowRect(activeHWnd, ref currentRect);
                int currentWidth = currentRect.Right - currentRect.Left;
                int currentHeight = currentRect.Bottom - currentRect.Top;

                var rect = nextScreen.WorkingArea;
                int width = move.WidthPercent.GetValueOrDefault(0) == 0 ? currentWidth : (rect.Width * move.WidthPercent.GetValueOrDefault(100) / 100);
                int height = move.HeightPercent.GetValueOrDefault(0) == 0 ? currentHeight : (rect.Height * move.HeightPercent.GetValueOrDefault(100) / 100);
                int x = rect.Left + (rect.Width - width) / 2;
                int y = rect.Top + (rect.Height - height) / 2;
                bool success = Win32ApiHelper.SetWindowPos(activeHWnd, IntPtr.Zero, x, y, width, height, 0x0040 /*SWP_SHOWWINDOW*/);
                if (!success)
                {
                    int error = System.Runtime.InteropServices.Marshal.GetLastWin32Error();
                    return $"Failed to move window to next monitor. Win32 error: {error}";
                }
                return "Window moved to other monitor.";
            }
            return "[Stub] Window move not implemented for: " + move.ToString();
        }
    }
}
