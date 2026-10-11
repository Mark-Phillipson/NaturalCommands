using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using NaturalCommands.Helpers;

namespace NaturalCommands
{
    /// <summary>
    /// Non-intrusive overlay form that displays a blue circular countdown indicator
    /// for the auto-click feature. Shows progress and remaining time near the cursor.
    /// </summary>
    public class AutoClickOverlayForm : Form
    {
        private static AutoClickOverlayForm? _instance;
        private static readonly object _lock = new object();
        private static System.Threading.SynchronizationContext? _uiContext;
        private static int _updatePending;
        private static int _uiThreadId;
        private float _percentage = 0;
        private int _remainingMs = 0;

        // Track last-painted state to avoid unnecessary repaints and reduce flicker
        private float _lastPaintedPercentage = -1f;
        private System.Drawing.Point _lastLocation = System.Drawing.Point.Empty;
        private const float _minUpdateDelta = 1f; // percent

        // Cycle/animation state: switch progress color each countdown cycle and alternate a background color when a cycle completes
        private static int _cycleIndex = 0;
        private Color _currentProgressColor = Color.FromArgb(230, 0, 200, 0); // default bright green
        private Color _currentBackgroundFill = Color.FromArgb(0, 0, 0, 0); // transparent by default
        private static readonly Color[] _progressColors = new[] {
            Color.FromArgb(230, 0, 200, 0),    // green
            Color.FromArgb(230, 255, 165, 0),  // orange
            Color.FromArgb(230, 0, 120, 215)   // blue
        };
        private static readonly Color[] _backgroundCycleColors = new[] {
            Color.FromArgb(120, 0, 200, 0),    // semi-transparent green
            Color.FromArgb(120, 255, 165, 0),  // semi-transparent orange
            Color.FromArgb(120, 0, 120, 215)   // semi-transparent blue
        };

        private const int WS_EX_TRANSPARENT = 0x20;
        private const int WS_EX_TOOLWINDOW = 0x80;
        private const int WS_EX_LAYERED = 0x80000;

        private AutoClickOverlayForm()
        {
            Logger.LogInfo(
$"[Overlay THREAD] Constructor: " +
$"ThreadId={Environment.CurrentManagedThreadId}, " +
$"SynchronizationContext={SynchronizationContext.Current?.GetType().FullName ?? "null"}");

            // Form properties for transparent, topmost overlay
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            BackColor = Color.Magenta;
            TransparencyKey = Color.Magenta;
            StartPosition = FormStartPosition.Manual;

            // Small circular overlay sized to replace the system cursor while counting down
            Width = 48;
            Height = 48;

            // Double buffering for smooth rendering
            DoubleBuffered = true;
            SetStyle(ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer, true);
        }


        protected override void OnPaint(PaintEventArgs e)
        {
            Logger.LogInfo(
    $"[Overlay PAINT] Percentage={_percentage:F1}, " +
    $"Location={Location}, Visible={Visible}");

            Graphics g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;

            // Keep the background colour consistent with TransparencyKey.
            g.Clear(Color.Magenta);

            int centerX = ClientSize.Width / 2;
            int centerY = ClientSize.Height / 2;
            int radius = Math.Min(ClientSize.Width, ClientSize.Height) / 2 - 5;

            // Define the circle's drawing area.
            int diameter = radius * 2;
            int x = centerX - radius;
            int y = centerY - radius;

            // Draw the full circle as the countdown background.
            using var backgroundPen = new Pen(Color.DimGray, 3);
            g.DrawEllipse(backgroundPen, x, y, diameter, diameter);

            // Draw the progress arc clockwise from the top.
            float progress = Math.Clamp(_percentage, 0f, 100f);
            float sweepAngle = progress / 100f * 360f;

            if (sweepAngle > 0)
            {
                using var progressPen = new Pen(_currentProgressColor, 4);
                progressPen.StartCap = LineCap.Round;
                progressPen.EndCap = LineCap.Round;

                g.DrawArc(
                    progressPen,
                    x, y, diameter, diameter,
                    -90, sweepAngle);
            }
        }


        /// <summary>
        /// Initialize the UI context - must be called from the UI thread before using the overlay.
        /// </summary>
        public static void InitializeUIContext()
        {
            if (SynchronizationContext.Current
                is not WindowsFormsSynchronizationContext)
            {
                SynchronizationContext.SetSynchronizationContext(
                    new WindowsFormsSynchronizationContext());
            }

            _uiContext = SynchronizationContext.Current;
            _uiThreadId = Environment.CurrentManagedThreadId;

            Logger.LogInfo(
                $"[Overlay THREAD] UI context initialized: " +
                $"ThreadId={_uiThreadId}, " +
                $"Context={_uiContext?.GetType().FullName ?? "null"}");
        }

        /// <summary>
        /// Shows or updates the overlay with current countdown information.
        /// </summary>
        public static void UpdateOverlay(Point cursorPos, int remainingMs, float percentage)
        {
            Logger.LogInfo(
    $"[Overlay THREAD] UpdateOverlay: " +
    $"ThreadId={Environment.CurrentManagedThreadId}, " +
    $"SynchronizationContext={SynchronizationContext.Current?.GetType().FullName ?? "null"}");

            Logger.LogInfo(
            $"[COUNTDOWN TEST] Cursor={cursorPos.X},{cursorPos.Y}, " +
            $"RemainingMs={remainingMs}, " +
            $"Percentage={percentage:F1}");
            if (!AutoClickManager.IsActive())
            {
                return;
            }

            // Check if we need to invoke on UI thread
            if (_uiContext != null &&
    Environment.CurrentManagedThreadId != _uiThreadId)
            {
                _uiContext.Post(_ =>
                {
                    UpdateOverlay(cursorPos, remainingMs, percentage);
                }, null);

                return;
            }

            lock (_lock)
            {
                try
                {
                    Logger.LogInfo("[Overlay DEBUG] UpdateOverlay entered.");

                    if (_instance == null || _instance.IsDisposed)
                    {
                        Logger.LogInfo("[Overlay DEBUG] Creating overlay form.");
                        _instance = new AutoClickOverlayForm();
                        Logger.LogInfo("[Overlay DEBUG] Overlay form created.");
                    }

                    Logger.LogInfo(
                        $"[Overlay DEBUG] Before Show: " +
                        $"Visible={_instance.Visible}, " +
                        $"HandleCreated={_instance.IsHandleCreated}, " +
                        $"Bounds={_instance.Bounds}");

                    _instance.Show();

                    Logger.LogInfo(
                        $"[Overlay DEBUG] After Show: " +
                        $"Visible={_instance.Visible}, " +
                        $"HandleCreated={_instance.IsHandleCreated}, " +
                        $"Bounds={_instance.Bounds}");

                    // Update position and values - overlay is anchored at the specified point
                    _instance.Location = new Point(cursorPos.X - _instance.Width / 2, cursorPos.Y - _instance.Height / 2);
                    _instance._remainingMs = remainingMs;
                    Logger.LogInfo(
                        $"[OVERLAY UPDATE TEST] Incoming={percentage:F1}, " +
                        $"Old={_instance._percentage:F1}, " +
                        $"IncomingPosition={cursorPos}, " +
                        $"Thread={Environment.CurrentManagedThreadId}");
                    // Detect start of a new countdown cycle (percentage has moved from 0 to >0)
                    bool startingCycle = _instance._percentage <= 0 && percentage > 0;
                    _instance._percentage = percentage;

                    if (startingCycle)
                    {
                        _instance.StartNewCycle();
                    }

                    // Keep system cursor visible so it can move independently of the overlay

                    if (!_instance.Visible)
                    {
                        _instance.Show();
                    }

                    if (!_instance.TopMost)
                    {
                        _instance.TopMost = true;
                    }

                    // Only invalidate if percentage or location changed enough to avoid excess repaints (reduces flicker)
                    // if (Math.Abs(_instance._percentage - _instance._lastPaintedPercentage) < _minUpdateDelta && _instance.Location == _instance._lastLocation)
                    // {
                    // }
                    // else
                    // {
                    //     _instance.Invalidate();
                    //     _instance._lastPaintedPercentage = _instance._percentage;
                    //     _instance._lastLocation = _instance.Location;
                    // }
                    _instance.Invalidate();
                    _instance._lastPaintedPercentage = _instance._percentage;
                    _instance._lastLocation = _instance.Location;
                    _instance.Update();
                }
                catch (Exception ex)
                {
                    NaturalCommands.Helpers.Logger.LogError($"[Overlay] ERROR updating overlay: {ex.Message}\n{ex.StackTrace}");
                }
            }
        }

        /// <summary>
        /// Called when a new countdown cycle starts so the overlay can change colors between cycles.
        /// </summary>
        private void StartNewCycle()
        {
            try
            {
                _cycleIndex = (_cycleIndex + 1) % _progressColors.Length;
                _currentProgressColor = _progressColors[_cycleIndex];
                // Choose the background color for when this cycle completes
                _currentBackgroundFill = _backgroundCycleColors[_cycleIndex];
            }
            catch (Exception ex)
            {
                NaturalCommands.Helpers.Logger.LogError($"[Overlay] StartNewCycle error: {ex.Message}");
            }
        }

        /// <summary>
        /// Hides the overlay.
        /// </summary>
        public static void HideOverlay()
        {
            // Check if we need to invoke on UI thread
            if (_uiContext != null && System.Threading.SynchronizationContext.Current != _uiContext)
            {
                _uiContext.Post(_ => HideOverlay(), null);
                return;
            }

            lock (_lock)
            {
                if (_instance != null && !_instance.IsDisposed)
                {
                    try
                    {

                        _instance.Hide();
                    }
                    catch (Exception ex)
                    {
                        NaturalCommands.Helpers.Logger.LogError($"Error hiding overlay: {ex.Message}");
                    }
                }
            }
        }

        /// <summary>
        /// Completely closes and disposes the overlay.
        /// </summary>
        public static void CloseOverlay()
        {
            // Check if we need to invoke on UI thread
            if (_uiContext != null && System.Threading.SynchronizationContext.Current != _uiContext)
            {
                _uiContext.Post(_ => CloseOverlay(), null);
                return;
            }

            lock (_lock)
            {
                if (_instance != null && !_instance.IsDisposed)
                {
                    try
                    {

                        _instance.Close();
                        _instance.Dispose();
                    }
                    catch (Exception ex)
                    {
                        NaturalCommands.Helpers.Logger.LogError($"Error closing overlay: {ex.Message}");
                    }
                    finally
                    {
                        _instance = null;
                    }
                }
            }
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
        }

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ExStyle |= WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_TOOLWINDOW;
                return cp;
            }
        }
    }
}
