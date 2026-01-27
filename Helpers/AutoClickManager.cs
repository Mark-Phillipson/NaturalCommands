using System;
using System.Runtime.InteropServices;
using System.Threading;
using NaturalCommands.Models;

namespace NaturalCommands.Helpers
{
    /// <summary>
    /// Manages automatic mouse clicking when the cursor is idle for a configured duration.
    /// Useful for hands-free gameplay and other voice-controlled scenarios.
    /// </summary>
    public static class AutoClickManager
    {
        // Win32 API imports
        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, int dwExtraInfo);

        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(POINT point);

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string? lpClassName, string? lpWindowName);

        [DllImport("user32.dll")]
        private static extern bool IsChild(IntPtr hWndParent, IntPtr hWnd);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        // Mouse event constants
        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;

        // State management
        private static System.Threading.Timer? _monitorTimer;
        private static POINT _lastPosition;
        private static int _idleTimeMs = 0;
        private static bool _isActive = false;
        private static readonly int _pollIntervalMs = 50; // Check mouse position every 50ms
        private static readonly int _movementTolerancePixels = 5; // Allow small movements within this radius (reduced from 100 to 5 pixels)

        /// <summary>
        /// Starts auto-click mode with the specified delay.
        /// </summary>
        /// <param name="delayMs">Delay in milliseconds before clicking (100-2000ms). If 0, uses setting from AppSettings.</param>
        public static string Start(int delayMs = 0)
        {
            if (_isActive)
            {
                return "Auto-click mode is already active";
            }

            // Get delay from settings or parameter
            if (delayMs == 0)
            {
                delayMs = AppSettings.Instance.AutoClick.DelayMs;
            }

            // Validate bounds
            if (delayMs < 100 || delayMs > 2000)
            {
                delayMs = Math.Clamp(delayMs, 100, 2000);
                Logger.LogWarning($"Auto-click delay out of bounds, clamped to {delayMs}ms");
            }

            // Update settings
            AppSettings.Instance.AutoClick.DelayMs = delayMs;
            AppSettings.Instance.AutoClick.Enabled = true;
            
            try
            {
                AppSettings.Instance.Save();
            }
            catch (Exception ex)
            {
                Logger.LogError($"Failed to save auto-click settings: {ex.Message}");
            }

            // Initialize state
            GetCursorPos(out _lastPosition);
            _idleTimeMs = 0;
            _isActive = true;

            // Start monitoring timer
            _monitorTimer = new System.Threading.Timer(
                MonitorTimer_Tick,
                null,
                0,                  // Start immediately
                _pollIntervalMs     // Check every 50ms
            );

            Logger.LogDebug($"[AutoClick] Timer created and started. Initial position: ({_lastPosition.X}, {_lastPosition.Y})");

            // Update tray icon to show active state
            TrayNotificationHelper.SetAutoClickActive(true);

            Logger.LogInfo($"Auto-click mode started with {delayMs}ms delay (tolerance: {_movementTolerancePixels}px)");
            Logger.LogDebug($"[AutoClick] ShowOverlay setting: {AppSettings.Instance.AutoClick.ShowOverlay}");
            return $"Auto-click enabled ({delayMs}ms delay, {_movementTolerancePixels}px tolerance)";
        }

        /// <summary>
        /// Stops auto-click mode.
        /// </summary>
        public static string Stop()
        {
            Logger.LogDebug("[AutoClick] Stop() method called");
            
            if (!_isActive)
            {
                Logger.LogDebug("[AutoClick] Stop() - not active, returning early");
                return "Auto-click mode is not active";
            }

            Logger.LogDebug("[AutoClick] Stop() - stopping timer");
            
            // Stop timer
            if (_monitorTimer != null)
            {
                _monitorTimer.Change(Timeout.Infinite, Timeout.Infinite);
                _monitorTimer.Dispose();
                _monitorTimer = null;
                Logger.LogDebug("[AutoClick] Stop() - timer disposed");
            }

            // Hide overlay
            Logger.LogDebug("[AutoClick] Stop() - hiding overlay");
            AutoClickOverlayForm.HideOverlay();

            // Update tray icon to show inactive state
            Logger.LogDebug("[AutoClick] Stop() - updating tray icon");
            TrayNotificationHelper.SetAutoClickActive(false);

            // Update settings
            AppSettings.Instance.AutoClick.Enabled = false;
            try
            {
                AppSettings.Instance.Save();
                Logger.LogDebug("[AutoClick] Stop() - settings saved");
            }
            catch (Exception ex)
            {
                Logger.LogError($"Failed to save auto-click settings: {ex.Message}");
            }

            _isActive = false;
            _idleTimeMs = 0;

            Logger.LogInfo("Auto-click mode stopped");
            Logger.LogDebug("[AutoClick] Stop() - completed successfully");
            return "Auto-click disabled";
        }

        /// <summary>
        /// Timer tick handler that monitors mouse position and triggers click when idle.
        /// </summary>
        private static void MonitorTimer_Tick(object? state)
        {
            if (!_isActive)
            {
                Logger.LogDebug("[AutoClick] Timer tick but not active, returning");
                return;
            }

            try
            {
                // Get current cursor position
                if (!GetCursorPos(out POINT currentPos))
                {
                    Logger.LogError("Failed to get cursor position in auto-click monitor");
                    return;
                }

                Logger.LogDebug($"[AutoClick] Tick - Cursor: ({currentPos.X}, {currentPos.Y}), Last: ({_lastPosition.X}, {_lastPosition.Y}), IdleTime: {_idleTimeMs}ms");

                // Check if mouse is over taskbar - if so, stop auto-click
                if (IsMouseOverTaskbar(currentPos))
                {
                    Logger.LogInfo("[AutoClick] Mouse is over taskbar, stopping auto-click");
                    Stop();
                    return;
                }

                // Check if mouse has moved beyond tolerance threshold
                int deltaX = Math.Abs(currentPos.X - _lastPosition.X);
                int deltaY = Math.Abs(currentPos.Y - _lastPosition.Y);
                double distance = Math.Sqrt(deltaX * deltaX + deltaY * deltaY);
                bool hasMovedSignificantly = distance > _movementTolerancePixels;

                Logger.LogDebug($"[AutoClick] Movement - DeltaX: {deltaX}, DeltaY: {deltaY}, Distance: {distance:F2}px, Threshold: {_movementTolerancePixels}px, Moved: {hasMovedSignificantly}");

                if (hasMovedSignificantly)
                {
                    // Mouse moved significantly - reset idle timer and update position
                    Logger.LogDebug($"[AutoClick] Mouse moved beyond tolerance, resetting idle time");
                    _lastPosition = currentPos;
                    _idleTimeMs = 0;
                    
                    // Hide overlay when mouse moves
                    Logger.LogDebug($"[AutoClick] Hiding overlay due to movement");
                    AutoClickOverlayForm.HideOverlay();
                }
                else
                {
                    // Mouse is idle - increment timer
                    _idleTimeMs += _pollIntervalMs;

                    int delayMs = AppSettings.Instance.AutoClick.DelayMs;
                    Logger.LogDebug($"[AutoClick] Mouse idle - IdleTime: {_idleTimeMs}ms / {delayMs}ms");
                    
                    // Show overlay if enabled and we're counting down
                    if (AppSettings.Instance.AutoClick.ShowOverlay && _idleTimeMs > 0)
                    {
                        int remainingMs = Math.Max(0, delayMs - _idleTimeMs);
                        float percentage = Math.Min(100, (_idleTimeMs / (float)delayMs) * 100);
                        
                        Logger.LogDebug($"[AutoClick] Updating overlay - Remaining: {remainingMs}ms, Percentage: {percentage:F1}%");
                        var point = new System.Drawing.Point(currentPos.X, currentPos.Y);
                        AutoClickOverlayForm.UpdateOverlay(point, remainingMs, percentage);
                    }
                    else if (!AppSettings.Instance.AutoClick.ShowOverlay)
                    {
                        Logger.LogDebug($"[AutoClick] Overlay disabled in settings");
                    }

                    // Check if idle time has reached the threshold
                    if (_idleTimeMs >= delayMs)
                    {
                        Logger.LogInfo($"[AutoClick] Idle threshold reached ({_idleTimeMs}ms >= {delayMs}ms), performing click");
                        // Perform the click
                        PerformClick(currentPos);
                        
                        // Reset idle timer to wait for next click
                        _idleTimeMs = 0;
                        
                        // Hide overlay after click
                        AutoClickOverlayForm.HideOverlay();
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"Error in auto-click monitor: {ex.Message}");
            }
        }

        /// <summary>
        /// Performs a left click at the current cursor position.
        /// </summary>
        private static void PerformClick(POINT pos)
        {
            try
            {
                mouse_event(MOUSEEVENTF_LEFTDOWN, pos.X, pos.Y, 0, 0);
                mouse_event(MOUSEEVENTF_LEFTUP, pos.X, pos.Y, 0, 0);
                
                Logger.LogDebug($"Auto-click performed at ({pos.X}, {pos.Y})");
            }
            catch (Exception ex)
            {
                Logger.LogError($"Error performing auto-click: {ex.Message}");
            }
        }

        /// <summary>
        /// Checks if the mouse cursor is over the Windows taskbar.
        /// </summary>
        private static bool IsMouseOverTaskbar(POINT point)
        {
            try
            {
                // Get the window at the cursor position
                IntPtr hWnd = WindowFromPoint(point);
                if (hWnd == IntPtr.Zero)
                    return false;

                // Find the taskbar window (Shell_TrayWnd is the taskbar class name)
                IntPtr taskbarHandle = FindWindow("Shell_TrayWnd", null);
                if (taskbarHandle == IntPtr.Zero)
                    return false;

                // Check if the window at cursor is the taskbar or a child of the taskbar
                if (hWnd == taskbarHandle || IsChild(taskbarHandle, hWnd))
                {
                    return true;
                }

                // Also check for secondary taskbars on multi-monitor setups (Shell_SecondaryTrayWnd)
                IntPtr secondaryTaskbarHandle = FindWindow("Shell_SecondaryTrayWnd", null);
                if (secondaryTaskbarHandle != IntPtr.Zero)
                {
                    if (hWnd == secondaryTaskbarHandle || IsChild(secondaryTaskbarHandle, hWnd))
                    {
                        return true;
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                Logger.LogError($"Error checking taskbar position: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Returns whether auto-click mode is currently active.
        /// </summary>
        public static bool IsActive()
        {
            return _isActive;
        }

        /// <summary>
        /// Gets the current auto-click delay in milliseconds.
        /// </summary>
        public static int GetDelayMs()
        {
            return AppSettings.Instance.AutoClick.DelayMs;
        }

        /// <summary>
        /// Sets the auto-click delay in milliseconds (100-2000ms).
        /// </summary>
        public static string SetDelay(int delayMs)
        {
            // Validate bounds
            if (delayMs < 100 || delayMs > 2000)
            {
                return $"Invalid delay {delayMs}ms. Must be between 100-2000ms.";
            }

            AppSettings.Instance.AutoClick.DelayMs = delayMs;
            
            try
            {
                AppSettings.Instance.Save();
                Logger.LogInfo($"Auto-click delay set to {delayMs}ms");
                return $"Auto-click delay set to {delayMs}ms";
            }
            catch (Exception ex)
            {
                Logger.LogError($"Failed to save auto-click delay: {ex.Message}");
                return $"Failed to save delay: {ex.Message}";
            }
        }
    }
}
