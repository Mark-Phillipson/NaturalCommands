using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.IO;
using System.IO.MemoryMappedFiles;

namespace NaturalCommands.Helpers
{
    /// <summary>
    /// Manages continuous mouse movement via voice commands.
    /// Supports 8-directional movement (cardinal + diagonal), variable speed, and click actions.
    /// </summary>
    public static class MouseMoveManager
    {
        // Win32 API imports
        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, int dwExtraInfo);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int X;
            public int Y;
        }

        // Mouse event constants
        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;
        private const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
        private const uint MOUSEEVENTF_RIGHTUP = 0x0010;

        // State management
        private static System.Threading.Timer? _moveTimer;
        private static string? _currentDirection;
        private static int _speed = 5; // pixels per tick - will be tuned during testing
        private static readonly int _minSpeed = 2;
        private static readonly int _maxSpeed = 50;
        
        // Inter-process communication using a named event and memory-mapped file
        private static EventWaitHandle? _stopSignal;
        private const string StopSignalName = "NaturalCommands_StopMouseMove";
        private const string SpeedMemoryName = "NaturalCommands_MouseSpeed";
        private static MemoryMappedFile? _speedMemory;

        /// <summary>
        /// Starts moving the mouse continuously in the specified direction.
        /// Supports cardinal (up, down, left, right) and diagonal (up left, up right, down left, down right) directions.
        /// </summary>
        public static string StartMoving(string direction)
        {
            // Normalize direction input
            direction = direction.ToLower().Trim();

            // Validate direction
            if (!IsValidDirection(direction))
            {
                return $"Invalid direction '{direction}'. Supported: up, down, left, right, up left, up right, down left, down right.";
            }

            // Stop any existing movement
            StopMoving(performClick: false);

            // Start new movement
            _currentDirection = direction;
            
            // Create or open the stop signal event
            try
            {
                _stopSignal = new EventWaitHandle(false, EventResetMode.ManualReset, StopSignalName);
                _stopSignal.Reset(); // Ensure it's not signaled
            }
            catch (Exception ex)
            {
                Logger.LogError($"Failed to create stop signal: {ex.Message}");
            }
            
            // Create or open shared memory for speed
            try
            {
                _speedMemory = MemoryMappedFile.CreateNew(SpeedMemoryName, sizeof(int), MemoryMappedFileAccess.ReadWrite);
                WriteSpeedToMemory(_speed);
                Logger.LogDebug($"Created speed memory with initial speed: {_speed}");
            }
            catch (IOException)
            {
                // File already exists, try to open it
                try
                {
                    _speedMemory = MemoryMappedFile.OpenExisting(SpeedMemoryName, MemoryMappedFileRights.ReadWrite);
                    WriteSpeedToMemory(_speed);
                    Logger.LogDebug($"Opened existing speed memory and updated to: {_speed}");
                }
                catch (Exception ex2)
                {
                    Logger.LogError($"Failed to open existing speed memory: {ex2.Message}");
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"Failed to create speed memory: {ex.Message}");
            }
            
            _moveTimer = new System.Threading.Timer(
                MoveTimer_Tick,
                null,
                0,              // Start immediately
                16              // ~60 FPS for smooth movement
            );

            Logger.LogInfo($"Started moving mouse {direction} at speed {_speed}");
            return $"Moving mouse {direction}";
        }

        /// <summary>
        /// Stops mouse movement and optionally performs a click action.
        /// </summary>
        public static string StopMoving(bool performClick = false, bool isRightClick = false)
        {
            // Signal other processes to stop
            try
            {
                using (var signal = EventWaitHandle.OpenExisting(StopSignalName))
                {
                    signal.Set();
                    Logger.LogInfo("Sent stop signal to running mouse movement process");
                }
            }
            catch (WaitHandleCannotBeOpenedException)
            {
                // No active movement process
                Logger.LogDebug("No active mouse movement process found");
            }
            catch (Exception ex)
            {
                Logger.LogError($"Failed to send stop signal: {ex.Message}");
            }
            
            // Also stop local timer if running in same process
            if (_moveTimer != null)
            {
                _moveTimer.Change(Timeout.Infinite, Timeout.Infinite); // Stop the timer
                _moveTimer.Dispose();
                _moveTimer = null;
            }
            
            if (_stopSignal != null)
            {
                _stopSignal.Dispose();
                _stopSignal = null;
            }
            
            if (_speedMemory != null)
            {
                _speedMemory.Dispose();
                _speedMemory = null;
            }

            string result = "Stopped mouse movement";

            if (performClick)
            {
                PerformClick(isRightClick);
                result = isRightClick ? "Stopped and right-clicked" : "Stopped and clicked";
                Logger.LogInfo(result);
            }
            else if (_currentDirection != null)
            {
                Logger.LogInfo("Stopped mouse movement");
            }

            _currentDirection = null;
            return result;
        }

        /// <summary>
        /// Adjusts the mouse movement speed.
        /// </summary>
        public static string AdjustSpeed(string adjustment)
        {
            adjustment = adjustment.ToLower().Trim();

            int oldSpeed = _speed;
            int newSpeed = oldSpeed;

            if (adjustment == "faster")
            {
                newSpeed = Math.Min(oldSpeed + 5, _maxSpeed);
            }
            else if (adjustment == "slower")
            {
                newSpeed = Math.Max(oldSpeed - 5, _minSpeed);
            }
            else
            {
                return $"Invalid speed adjustment '{adjustment}'. Use 'faster' or 'slower'.";
            }

            // Update local speed
            _speed = newSpeed;
            
            // Try to update the shared memory for inter-process communication
            try
            {
                using (var mmf = MemoryMappedFile.OpenExisting(SpeedMemoryName, MemoryMappedFileRights.ReadWrite))
                {
                    WriteSpeedToMemory(newSpeed, mmf);
                    Logger.LogInfo($"Speed adjusted from {oldSpeed} to {newSpeed} (shared with other process)");
                }
            }
            catch (FileNotFoundException)
            {
                // No active mouse movement process, just update local
                Logger.LogInfo($"Speed adjusted from {oldSpeed} to {newSpeed} (local only - no active movement)");
            }
            catch (Exception ex)
            {
                Logger.LogError($"Failed to update shared speed: {ex.Message}");
                Logger.LogInfo($"Speed adjusted from {oldSpeed} to {newSpeed} (local only)");
            }

            string result = $"Speed adjusted from {oldSpeed} to {newSpeed}";
            return result;
        }

        /// <summary>
        /// Timer tick handler that moves the mouse based on current direction and speed.
        /// </summary>
        private static void MoveTimer_Tick(object? state)
        {
            if (string.IsNullOrEmpty(_currentDirection))
            {
                return;
            }
            
            // Read the current speed from shared memory
            try
            {
                int sharedSpeed = ReadSpeedFromMemory();
                if (sharedSpeed != _speed && sharedSpeed >= _minSpeed && sharedSpeed <= _maxSpeed)
                {
                    _speed = sharedSpeed;
                }
            }
            catch
            {
                // If we can't read shared memory, just use local speed
            }

            try
            {
                // Get current cursor position
                if (!GetCursorPos(out POINT currentPos))
                {
                    Logger.LogError("Failed to get cursor position");
                    return;
                }

                // Calculate new position based on direction and speed
                int deltaX = 0;
                int deltaY = 0;

                switch (_currentDirection)
                {
                    case "up":
                        deltaY = -_speed;
                        break;
                    case "down":
                        deltaY = _speed;
                        break;
                    case "left":
                        deltaX = -_speed;
                        break;
                    case "right":
                        deltaX = _speed;
                        break;
                    case "up left":
                    case "left up":
                        deltaX = -_speed;
                        deltaY = -_speed;
                        break;
                    case "up right":
                    case "right up":
                        deltaX = _speed;
                        deltaY = -_speed;
                        break;
                    case "down left":
                    case "left down":
                        deltaX = -_speed;
                        deltaY = _speed;
                        break;
                    case "down right":
                    case "right down":
                        deltaX = _speed;
                        deltaY = _speed;
                        break;
                }

                // Update cursor position
                int newX = currentPos.X + deltaX;
                int newY = currentPos.Y + deltaY;

                SetCursorPos(newX, newY);
            }
            catch (Exception ex)
            {
                Logger.LogError($"Error moving mouse: {ex.Message}");
            }
        }

        /// <summary>
        /// Performs a mouse click at the current cursor position.
        /// </summary>
        private static void PerformClick(bool isRightClick)
        {
            try
            {
                // Get current cursor position for click
                if (!GetCursorPos(out POINT pos))
                {
                    Logger.LogError("Failed to get cursor position for click");
                    return;
                }

                if (isRightClick)
                {
                    mouse_event(MOUSEEVENTF_RIGHTDOWN, pos.X, pos.Y, 0, 0);
                    mouse_event(MOUSEEVENTF_RIGHTUP, pos.X, pos.Y, 0, 0);
                }
                else
                {
                    mouse_event(MOUSEEVENTF_LEFTDOWN, pos.X, pos.Y, 0, 0);
                    mouse_event(MOUSEEVENTF_LEFTUP, pos.X, pos.Y, 0, 0);
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"Error performing click: {ex.Message}");
            }
        }

        /// <summary>
        /// Validates if the direction string is supported.
        /// </summary>
        private static bool IsValidDirection(string direction)
        {
            var validDirections = new[]
            {
                "up", "down", "left", "right",
                "up left", "left up", "up right", "right up",
                "down left", "left down", "down right", "right down"
            };

            return Array.Exists(validDirections, d => d.Equals(direction, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// Returns whether mouse movement is currently active.
        /// </summary>
        public static bool IsMoving()
        {
            return _moveTimer != null && !string.IsNullOrEmpty(_currentDirection);
        }
        
        /// <summary>
        /// Writes the current speed to shared memory.
        /// </summary>
        private static void WriteSpeedToMemory(int speed, MemoryMappedFile? mmf = null)
        {
            try
            {
                var file = mmf ?? _speedMemory;
                if (file != null)
                {
                    using (var accessor = file.CreateViewAccessor(0, sizeof(int)))
                    {
                        accessor.Write(0, speed);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"Error writing speed to memory: {ex.Message}");
            }
        }
        
        /// <summary>
        /// Reads the current speed from shared memory.
        /// </summary>
        private static int ReadSpeedFromMemory()
        {
            try
            {
                if (_speedMemory != null)
                {
                    using (var accessor = _speedMemory.CreateViewAccessor(0, sizeof(int)))
                    {
                        return accessor.ReadInt32(0);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"Error reading speed from memory: {ex.Message}");
            }
            return _speed;
        }
    }
}
