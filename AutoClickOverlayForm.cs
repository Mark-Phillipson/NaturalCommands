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
        // Track whether we've hidden the system cursor so we can restore it later
        private bool _cursorHidden = false;

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

            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

            // Calculate center and radius
            int centerX = Width / 2;
            int centerY = Height / 2;
            int radius = Math.Min(Width, Height) / 2 - 5;

            // Draw semi-transparent background circle
            using (var bgBrush = new SolidBrush(Color.FromArgb(160, 20, 20, 20)))
            {
                g.FillEllipse(bgBrush, centerX - radius, centerY - radius, radius * 2, radius * 2);
            }

            // Draw darker progress arc (more visible from periphery)
            if (_percentage > 0)
            {
                using (var progressPen = new Pen(Color.FromArgb(220, 180, 60, 0), 6))
                {
                    // Start at top (-90 degrees) and sweep clockwise
                    float sweepAngle = (_percentage / 100f) * 360f;
                    g.DrawArc(progressPen,
                        centerX - radius + 3,
                        centerY - radius + 3,
                        (radius - 3) * 2,
                        (radius - 3) * 2,
                        -90, // Start at top
                        sweepAngle);
                }

                // Draw a small center dot to give a clear 'cursor hotspot' visual when replacing the cursor
                using (var dotBrush = new SolidBrush(Color.White))
                {
                    int dotRadius = Math.Max(2, radius / 6);
                    g.FillEllipse(dotBrush, centerX - dotRadius, centerY - dotRadius, dotRadius * 2, dotRadius * 2);
                }
            }
        }

        /// <summary>
        /// Initialize the UI context - must be called from the UI thread before using the overlay.
        /// </summary>
        public static void InitializeUIContext()
        {
            _uiContext = System.Threading.SynchronizationContext.Current;
            NaturalCommands.Helpers.Logger.LogDebug($"[Overlay] UI context initialized: {_uiContext?.GetType().Name ?? "null"}");
        }

        /// <summary>
        /// Shows or updates the overlay with current countdown information.
        /// </summary>
        public static void UpdateOverlay(Point cursorPos, int remainingMs, float percentage)
        {
            NaturalCommands.Helpers.Logger.LogDebug($"[Overlay] UpdateOverlay called - Pos: ({cursorPos.X}, {cursorPos.Y}), Remaining: {remainingMs}ms, Pct: {percentage:F1}%");
            
            // Check if we need to invoke on UI thread
            if (_uiContext != null && System.Threading.SynchronizationContext.Current != _uiContext)
            {
                NaturalCommands.Helpers.Logger.LogDebug($"[Overlay] Posting to UI thread via captured SynchronizationContext");
                _uiContext.Post(_ => UpdateOverlay(cursorPos, remainingMs, percentage), null);
                return;
            }

            NaturalCommands.Helpers.Logger.LogDebug($"[Overlay] On UI thread - executing directly");

            lock (_lock)
            {
                try
                {
                    if (_instance == null || _instance.IsDisposed)
                    {
                        NaturalCommands.Helpers.Logger.LogDebug($"[Overlay] Creating new overlay form instance");
                        _instance = new AutoClickOverlayForm();
                        _instance.Show();
                        NaturalCommands.Helpers.Logger.LogDebug($"[Overlay] New form shown - Visible: {_instance.Visible}, Handle: {_instance.Handle}");
                    }

                    // Update position and values - center overlay on the cursor to visually replace it
                    _instance.Location = new Point(cursorPos.X - _instance.Width / 2, cursorPos.Y - _instance.Height / 2);
                    _instance._remainingMs = remainingMs;
                    _instance._percentage = percentage;

                    // Replace the system cursor while countdown is active
                    if (_instance._percentage > 0)
                    {
                        if (!_instance._cursorHidden)
                        {
                            Cursor.Hide();
                            _instance._cursorHidden = true;
                            NaturalCommands.Helpers.Logger.LogDebug("[Overlay] Cursor hidden to replace system cursor during countdown");
                        }
                    }
                    else
                    {
                        if (_instance._cursorHidden)
                        {
                            Cursor.Show();
                            _instance._cursorHidden = false;
                            NaturalCommands.Helpers.Logger.LogDebug("[Overlay] Cursor restored");
                        }
                    }
                    
                    NaturalCommands.Helpers.Logger.LogDebug($"[Overlay] Updated - Location: ({_instance.Location.X}, {_instance.Location.Y}), Visible: {_instance.Visible}, TopMost: {_instance.TopMost}");
                    
                    if (!_instance.Visible)
                    {
                        NaturalCommands.Helpers.Logger.LogDebug($"[Overlay] Form not visible, calling Show()");
                        _instance.Show();
                    }
                    
                    if (!_instance.TopMost)
                    {
                        NaturalCommands.Helpers.Logger.LogDebug($"[Overlay] Setting TopMost = true");
                        _instance.TopMost = true;
                    }
                    
                    // Force a refresh
                    _instance.Invalidate();
                    NaturalCommands.Helpers.Logger.LogDebug($"[Overlay] Invalidate called");
                }
                catch (Exception ex)
                {
                    NaturalCommands.Helpers.Logger.LogError($"[Overlay] ERROR updating overlay: {ex.Message}\n{ex.StackTrace}");
                }
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
                        // Restore cursor if we had hidden it
                        if (_instance._cursorHidden)
                        {
                            Cursor.Show();
                            _instance._cursorHidden = false;
                            NaturalCommands.Helpers.Logger.LogDebug("[Overlay] Cursor restored during HideOverlay");
                        }

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
                        // Ensure cursor is restored
                        if (_instance._cursorHidden)
                        {
                            Cursor.Show();
                            _instance._cursorHidden = false;
                            NaturalCommands.Helpers.Logger.LogDebug("[Overlay] Cursor restored during CloseOverlay");
                        }

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
