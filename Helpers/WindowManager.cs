using System;
using System.IO;
using System.Windows.Forms;
namespace NaturalCommands.Helpers
{
    // Handles window management actions (maximize, move, always on top, etc.)
    public class WindowManager
    {
        public static string ExecuteMoveWindow(MoveWindowAction move)
        {
            // Get active window handle
            IntPtr hWnd = NaturalCommands.Commands.GetForegroundWindow();
            // Maximize logic
            if ((move.Position == "center" || move.Position == null) && move.WidthPercent == 100 && move.HeightPercent == 100 && move.Monitor != "next")
            {
                Win32ApiHelper.SetForegroundWindow(hWnd);
                int style = Win32ApiHelper.GetWindowLong(hWnd, Win32ApiHelper.GWL_STYLE);
                bool canMaximize = (style & Win32ApiHelper.WS_MAXIMIZEBOX) != 0;
                var className = new System.Text.StringBuilder(256);
                Win32ApiHelper.GetClassName(hWnd, className, className.Capacity);
                if (!canMaximize)
                {
                    return "Window cannot be maximized (missing maximize button).";
                }
                const int SW_MAXIMIZE = 3;
                bool success = Win32ApiHelper.ShowWindow(hWnd, SW_MAXIMIZE);
                if (!success)
                {
                    int error = System.Runtime.InteropServices.Marshal.GetLastWin32Error();
                    return $"Failed to maximize window. Win32 error: {error}";
                }
                return "Window maximized.";
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
