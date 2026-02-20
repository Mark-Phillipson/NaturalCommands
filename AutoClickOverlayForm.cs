using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

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

        // Win32 API to make form click-through
        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        private const int GWL_EXSTYLE = -20;
        private const int WS_EX_LAYERED = 0x80000;
        private const int WS_EX_TRANSPARENT = 0x20;

        private AutoClickOverlayForm()
        {
            // Form properties for transparent, topmost overlay
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            BackColor = Color.Lime; // Transparency key color
            TransparencyKey = Color.Lime;
            StartPosition = FormStartPosition.Manual;
            
            // Small circular overlay sized to replace the system cursor while counting down
            Width = 48;
            Height = 48;
            
            // Double buffering for smooth rendering
            DoubleBuffered = true;
            SetStyle(ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.UserPaint |
                     ControlStyles.OptimizedDoubleBuffer, true);

            // Make form click-through (mouse events pass through)
            Load += (s, e) =>
            {
                int exStyle = GetWindowLong(Handle, GWL_EXSTYLE);
                SetWindowLong(Handle, GWL_EXSTYLE, exStyle | WS_EX_LAYERED | WS_EX_TRANSPARENT);
            };
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            // Use BufferedGraphics to reduce flicker
            var context = BufferedGraphicsManager.Current;
            using (var buffer = context.Allocate(e.Graphics, this.ClientRectangle))
            {
                var bg = buffer.Graphics;
                bg.SmoothingMode = SmoothingMode.AntiAlias;
                bg.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

                // Clear buffer with the form BackColor (the transparency key) to ensure per-pixel transparency
                // This avoids leftover black pixels when using BufferedGraphics + TransparencyKey.
                bg.Clear(this.BackColor);
                bg.CompositingMode = CompositingMode.SourceOver;

                // Calculate center and radius
                int centerX = Width / 2;
                int centerY = Height / 2;
                int radius = Math.Min(Width, Height) / 2 - 5;

                // Draw outer ring (transparent center) with a neutral dim color so progress stands out
                using (var ringPen = new Pen(Color.FromArgb(140, 40, 40, 40), 3))
                {
                    bg.DrawEllipse(ringPen, centerX - radius + 3, centerY - radius + 3, (radius - 3) * 2, (radius - 3) * 2);
                }

                // If the cycle has completed (full) draw a subtle filled background to indicate completion
                if (_percentage >= 100 && _currentBackgroundFill.A > 0)
                {
                    using (var fillBrush = new SolidBrush(_currentBackgroundFill))
                    {
                        bg.FillEllipse(fillBrush, centerX - radius + 6, centerY - radius + 6, (radius - 6) * 2, (radius - 6) * 2);
                    }
                }

                // Draw progress arc (bright accent) to clearly highlight progress
                if (_percentage > 0)
                {
                    using (var progressPen = new Pen(_currentProgressColor, 4))
                    {
                        float sweepAngle = (_percentage / 100f) * 360f;
                        bg.DrawArc(progressPen,
                            centerX - radius + 3,
                            centerY - radius + 3,
                            (radius - 3) * 2,
                            (radius - 3) * 2,
                            -90,
                            sweepAngle);
                    }
                }

                buffer.Render(e.Graphics);
            }
        }

        /// <summary>
        /// Initialize the UI context - must be called from the UI thread before using the overlay.
        /// </summary>
        public static void InitializeUIContext()
        {
            _uiContext = System.Threading.SynchronizationContext.Current;
        }

        /// <summary>
        /// Shows or updates the overlay with current countdown information.
        /// </summary>
        public static void UpdateOverlay(Point cursorPos, int remainingMs, float percentage)
        {
            // Check if we need to invoke on UI thread
            if (_uiContext != null && System.Threading.SynchronizationContext.Current != _uiContext)
            {
                _uiContext.Post(_ => UpdateOverlay(cursorPos, remainingMs, percentage), null);
                return;
            }

            lock (_lock)
            {
                try
                {
                    if (_instance == null || _instance.IsDisposed)
                    {
                        _instance = new AutoClickOverlayForm();
                        _instance.Show();
                    }

                    // Update position and values - overlay is anchored at the specified point
                    _instance.Location = new Point(cursorPos.X - _instance.Width / 2, cursorPos.Y - _instance.Height / 2);
                    _instance._remainingMs = remainingMs;

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
                    if (Math.Abs(_instance._percentage - _instance._lastPaintedPercentage) < _minUpdateDelta && _instance.Location == _instance._lastLocation)
                    {
                    }
                    else
                    {
                        _instance.Invalidate();
                        _instance._lastPaintedPercentage = _instance._percentage;
                        _instance._lastLocation = _instance.Location;
                    }
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
                cp.ExStyle |= WS_EX_LAYERED | WS_EX_TRANSPARENT;
                return cp;
            }
        }
    }
}
