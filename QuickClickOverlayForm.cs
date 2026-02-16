using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using NaturalCommands.Helpers;
using NaturalCommands.Models;

namespace NaturalCommands
{
    /// <summary>
    /// Overlay that displays Quick Click regions for a target application window.
    /// Supports a click-through display mode and an interactive edit mode.
    /// This is a scaffold implementing core behaviors (render, select, drag, basic edit UI).
    /// </summary>
    public class QuickClickOverlayForm : Form
    {
        private static QuickClickOverlayForm? _instance;
        private static readonly object _lock = new object();
        private static System.Threading.SynchronizationContext? _uiContext;

        private QuickClickProfile? _profile;
        private IntPtr _targetWindow = IntPtr.Zero;
        private Rectangle _targetWindowRect = Rectangle.Empty;

        private System.Windows.Forms.Timer _repositionTimer;
        private bool _editMode = false;

        // Interaction state
        private QuickClickRegion? _selectedRegion;
        private Point _dragOffset;
        private bool _isDragging = false;
        private bool _placementMode = false; // + New pressed

        // UI metrics
        private readonly Font _labelFont = new Font("Segoe UI", 10, FontStyle.Bold);
        private readonly Brush _regionFill = new SolidBrush(Color.FromArgb(140, 30, 144, 255));
        private readonly Brush _regionBorder = new SolidBrush(Color.FromArgb(200, 10, 10, 10));
        private readonly Brush _textBrush = Brushes.White;
        private readonly Pen _selectedPen = new Pen(Color.Yellow, 3);

        // Buttons layout (rendered in edit mode)
        private readonly List<(string Key, Rectangle Rect)> _editorButtons = new List<(string, Rectangle)>();

        // Placement preview
        private Point _placementPoint;

        private QuickClickOverlayForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            BackColor = Color.Lime;
            TransparencyKey = Color.Lime;
            StartPosition = FormStartPosition.Manual;
            DoubleBuffered = true;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true);

            Width = SystemInformation.VirtualScreen.Width;
            Height = SystemInformation.VirtualScreen.Height;
            Location = new Point(SystemInformation.VirtualScreen.Left, SystemInformation.VirtualScreen.Top);

            // Timer to poll window location so overlay follows the target window
            _repositionTimer = new System.Windows.Forms.Timer { Interval = 200 };
            _repositionTimer.Tick += (s, e) => RefreshPositionAndInvalidate();

            // Mouse handlers for edit interactions
            MouseDown += QuickClickOverlayForm_MouseDown;
            MouseMove += QuickClickOverlayForm_MouseMove;
            MouseUp += QuickClickOverlayForm_MouseUp;
            MouseClick += QuickClickOverlayForm_MouseClick;
            KeyDown += QuickClickOverlayForm_KeyDown;

            // Make sure the overlay doesn't take focus when shown
            Shown += (s, e) => { try { Win32ApiHelper.SetWindowPos(Handle, new IntPtr(-2), 0, 0, 0, 0, 0x0001 | 0x0002); } catch { } };
        }

        #region Public API (thread-safe)
        public static void InitializeUIContext()
        {
            _uiContext = System.Threading.SynchronizationContext.Current;
            Logger.LogDebug($"QuickClickOverlayForm: UI context initialized: {_uiContext?.GetType().Name ?? "null"}");
        }

        public static void ShowOverlay(QuickClickProfile? profile, IntPtr windowHandle)
        {
            if (_uiContext != null && System.Threading.SynchronizationContext.Current != _uiContext)
            {
                _uiContext.Post(_ => ShowOverlay(profile, windowHandle), null);
                return;
            }

            lock (_lock)
            {
                if (_instance == null || _instance.IsDisposed)
                {
                    _instance = new QuickClickOverlayForm();
                }

                _instance.InternalShow(profile, windowHandle);
            }
        }

        public static void HideOverlay()
        {
            if (_uiContext != null && System.Threading.SynchronizationContext.Current != _uiContext)
            {
                _uiContext.Post(_ => HideOverlay(), null);
                return;
            }

            lock (_lock)
            {
                if (_instance != null && !_instance.IsDisposed)
                {
                    _instance.InternalHide();
                }
            }
        }

        public static bool IsVisible => _instance != null && !_instance.IsDisposed && _instance.Visible;

        public static void EnterEditMode()
        {
            if (_uiContext != null && System.Threading.SynchronizationContext.Current != _uiContext)
            {
                _uiContext.Post(_ => EnterEditMode(), null);
                return;
            }

            lock (_lock)
            {
                if (_instance != null && !_instance.IsDisposed)
                {
                    _instance.ToggleEditMode(true);
                }
            }
        }

        /// <summary>
        /// Show the overlay for the current foreground window and enter placement/edit mode
        /// with the placement cursor positioned at the current mouse location. Creates a new
        /// in-memory profile when none exists so the user can add a region immediately.
        /// </summary>
        public static void BeginPlacementAtCursorForForegroundWindow()
        {
            if (_uiContext != null && System.Threading.SynchronizationContext.Current != _uiContext)
            {
                _uiContext.Post(_ => BeginPlacementAtCursorForForegroundWindow(), null);
                return;
            }

            lock (_lock)
            {
                if (_instance == null || _instance.IsDisposed)
                {
                    _instance = new QuickClickOverlayForm();
                }

                IntPtr hwnd = Helpers.Win32ApiHelper.GetForegroundWindow();
                string? proc = CurrentApplicationHelper.GetCurrentProcessName();
                string? title = CurrentApplicationHelper.GetCurrentWindowTitle();

                int? mw = null, mh = null;
                try
                {
                    var monitor = Win32ApiHelper.MonitorFromWindow(hwnd, 2);
                    if (monitor != IntPtr.Zero)
                    {
                        var info = new Win32ApiHelper.MONITORINFOEX();
                        info.cbSize = System.Runtime.InteropServices.Marshal.SizeOf(typeof(Win32ApiHelper.MONITORINFOEX));
                        if (Win32ApiHelper.GetMonitorInfo(monitor, ref info))
                        {
                            mw = info.rcMonitor.Right - info.rcMonitor.Left;
                            mh = info.rcMonitor.Bottom - info.rcMonitor.Top;
                        }
                    }
                }
                catch { }

                var profile = Helpers.QuickClickLoader.GetProfilesForApp(proc ?? string.Empty, title, mw, mh).FirstOrDefault()
                              ?? new QuickClickProfile { ProcessName = proc ?? string.Empty, WindowTitlePattern = title, MonitorWidth = mw, MonitorHeight = mh };

                _instance.InternalShow(profile, hwnd);
                _instance.ToggleEditMode(true);
                _instance._placementMode = true;
                _instance._placementPoint = System.Windows.Forms.Cursor.Position;
                _instance.Invalidate();
            }
        }
        #endregion

        #region Internal lifecycle
        private void InternalShow(QuickClickProfile? profile, IntPtr windowHandle)
        {
            _profile = profile;
            _targetWindow = windowHandle;
            RefreshPositionAndInvalidate();

            // Display mode should be click-through
            EnsureClickThrough(true);
            _repositionTimer.Start();

            if (!Visible)
                Show();
            BringToFront();
        }

        private void InternalHide()
        {
            _repositionTimer.Stop();
            _placementMode = false;
            _isDragging = false;
            _selectedRegion = null;
            ToggleEditMode(false);
            Hide();
        }
        #endregion

        #region Painting
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            var g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            // Draw warning if no profile provided
            if (_profile == null)
            {
                var msg = "No Quick Click profile for this app/monitor. Open editor to create one.";
                var size = g.MeasureString(msg, _labelFont);
                var rect = new RectangleF(20, 20, size.Width + 16, size.Height + 10);
                using (var bg = new SolidBrush(Color.FromArgb(200, 0, 0, 0))) g.FillRectangle(bg, rect);
                g.DrawString(msg, _labelFont, Brushes.White, rect.X + 8, rect.Y + 4);
                return;
            }

            // Draw regions
            int alpha = (int)(AppSettings.Instance.QuickClicks.OverlayOpacity * 255);
            foreach (var r in _profile.Regions)
            {
                var rect = RegionToScreenRect(r);
                if (rect.Width <= 0 || rect.Height <= 0) continue;

                using (var fill = new SolidBrush(Color.FromArgb(alpha / 2, 30, 144, 255)))
                using (var border = new Pen(Color.FromArgb(alpha, 10, 10, 10), 2))
                {
                    g.FillRectangle(fill, rect);
                    g.DrawRectangle(border, rect);
                }

                // Draw name + click-type badge
                var label = r.Name;
                var badge = ClickTypeBadge(r.ClickType);
                var text = string.IsNullOrWhiteSpace(badge) ? label : $"{label} [{badge}]";
                var textSize = g.MeasureString(text, _labelFont);
                var textRect = new RectangleF(rect.X + 6, rect.Y + 6, textSize.Width, textSize.Height);
                using (var bg = new SolidBrush(Color.FromArgb(180, 0, 0, 0))) g.FillRectangle(bg, textRect);
                g.DrawString(text, _labelFont, _textBrush, textRect.Location);

                // Selected highlight
                if (_selectedRegion == r)
                {
                    g.DrawRectangle(_selectedPen, rect.X - 2, rect.Y - 2, rect.Width + 4, rect.Height + 4);
                }
            }

            if (_editMode)
            {
                RenderEditorUI(g);
            }

            // Placement preview
            if (_placementMode)
            {
                var size = AppSettings.Instance.QuickClicks.DefaultRegionSize;
                var pr = new Rectangle(_placementPoint.X - size / 2, _placementPoint.Y - size / 2, size, size);
                using (var pen = new Pen(Color.Lime, 2)) g.DrawRectangle(pen, pr);
            }
        }

        private void RenderEditorUI(Graphics g)
        {
            // Simple toolbar in the top-left corner
            int x = 16; int y = 16; int h = 28; int gap = 8;
            _editorButtons.Clear();

            void DrawBtn(string label, Color bg)
            {
                var size = g.MeasureString(label, _labelFont);
                var rect = new Rectangle(x, y, (int)size.Width + 20, h);
                using (var b = new SolidBrush(bg)) g.FillRectangle(b, rect);
                g.DrawRectangle(Pens.Black, rect);
                g.DrawString(label, _labelFont, Brushes.White, rect.X + 10, rect.Y + 4);
                _editorButtons.Add((label, rect));
                x += rect.Width + gap;
            }

            DrawBtn("+ New", Color.FromArgb(90, 50, 150));
            DrawBtn("+ Rectangle", Color.FromArgb(70, 120, 70));
            DrawBtn("Delete", Color.FromArgb(180, 60, 60));
            DrawBtn("Save & Exit", Color.FromArgb(40, 120, 180));

            // If a region is selected, draw click-type selector nearby
            if (_selectedRegion != null)
            {
                var badgeX = 16; var badgeY = y + h + 12;
                var types = new[] { "Left", "Double", "Right" };
                foreach (var t in types)
                {
                    var rect = new Rectangle(badgeX, badgeY, 80, 24);
                    var isActive = string.Equals(_selectedRegion.ClickType.ToString(), t, StringComparison.OrdinalIgnoreCase);
                    using (var b = new SolidBrush(isActive ? Color.DodgerBlue : Color.FromArgb(80, 80, 80))) g.FillRectangle(b, rect);
                    g.DrawRectangle(Pens.Black, rect);
                    g.DrawString(t, SystemFonts.DefaultFont, Brushes.White, rect.X + 8, rect.Y + 3);
                    _editorButtons.Add((t, rect));
                    badgeX += rect.Width + 8;
                }
            }
        }

        private string ClickTypeBadge(QuickClickClickType type)
        {
            return type switch
            {
                QuickClickClickType.Left => "L",
                QuickClickClickType.Double => "2×",
                QuickClickClickType.Right => "R",
                QuickClickClickType.Middle => "M",
                _ => "L",
            };
        }
        #endregion

        #region Interaction handlers
        private void QuickClickOverlayForm_MouseClick(object? sender, MouseEventArgs e)
        {
            if (!_editMode)
                return;

            // If in placement mode, prompt for name/click-type and create the region at the clicked point
            if (_placementMode)
            {
                var wp = ScreenToWindowClient(e.Location);
                var size = AppSettings.Instance.QuickClicks.DefaultRegionSize;

                // Default click type from settings
                var defaultClick = Enum.TryParse<QuickClickClickType>(AppSettings.Instance.QuickClicks.DefaultClickType, true, out var dct) ? dct : QuickClickClickType.Left;

                if (CreateQuickClickForm.ShowDialog(out var regionName, out var chosenClickType, out var applyToMonitorOnly, e.Location, "New Region", defaultClick, true))
                {
                    var region = new QuickClickRegion
                    {
                        Name = string.IsNullOrWhiteSpace(regionName) ? "New Region" : regionName,
                        X = wp.X - size / 2,
                        Y = wp.Y - size / 2,
                        Width = size,
                        Height = size,
                        ClickType = chosenClickType
                    };

                    // If the profile doesn't have monitor info and user requested monitor-scoped, capture current monitor size
                    if (applyToMonitorOnly && (_profile?.MonitorWidth == null || _profile?.MonitorHeight == null))
                    {
                        try
                        {
                            var scr = Screen.FromPoint(e.Location).Bounds;
                            if (_profile != null)
                            {
                                _profile.MonitorWidth = scr.Width;
                                _profile.MonitorHeight = scr.Height;
                            }
                        }
                        catch { }
                    }

                    _profile!.Regions.Add(region);
                    _placementMode = false;
                    _selectedRegion = region;
                    Invalidate();
                }
                else
                {
                    // cancelled — leave placement mode
                    _placementMode = false;
                    Invalidate();
                }

                return;
            }

            // Check editor buttons first
            foreach (var btn in _editorButtons)
            {
                if (btn.Rect.Contains(e.Location))
                {
                    HandleEditorButton(btn.Key);
                    return;
                }
            }

            // Hit test regions (top-most first)
            var hit = _profile!.Regions.LastOrDefault(r => RegionToScreenRect(r).Contains(e.Location));
            if (hit != null)
            {
                _selectedRegion = hit;
                Invalidate();
            }
            else
            {
                _selectedRegion = null;
                Invalidate();
            }
        }

        private void QuickClickOverlayForm_MouseDown(object? sender, MouseEventArgs e)
        {
            if (!_editMode) return;
            if (_selectedRegion == null) return;

            var rect = RegionToScreenRect(_selectedRegion);
            if (rect.Contains(e.Location))
            {
                _isDragging = true;
                _dragOffset = new Point(e.Location.X - rect.X, e.Location.Y - rect.Y);
            }
        }

        private void QuickClickOverlayForm_MouseMove(object? sender, MouseEventArgs e)
        {
            if (!_editMode) return;
            if (_isDragging && _selectedRegion != null)
            {
                var windowPoint = ScreenToWindowClient(e.Location);
                _selectedRegion.X = windowPoint.X - _dragOffset.X;
                _selectedRegion.Y = windowPoint.Y - _dragOffset.Y;
                Invalidate();
            }
            else if (_placementMode)
            {
                _placementPoint = e.Location;
                Invalidate();
            }
        }

        private void QuickClickOverlayForm_MouseUp(object? sender, MouseEventArgs e)
        {
            if (!_editMode) return;
            _isDragging = false;
        }

        private void QuickClickOverlayForm_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape && _editMode)
            {
                ToggleEditMode(false);
                e.Handled = true;
            }
        }

        private void HandleEditorButton(string key)
        {
            switch (key)
            {
                case "+ New":
                    _placementMode = true;
                    break;
                case "+ Rectangle":
                    // For scaffold: treat as New for now
                    _placementMode = true;
                    break;
                case "Delete":
                    if (_selectedRegion != null && _profile != null)
                    {
                        _profile.Regions.Remove(_selectedRegion);
                        _selectedRegion = null;
                        Invalidate();
                    }
                    break;
                case "Save & Exit":
                    if (_profile != null)
                    {
                        QuickClickLoader.Save(new[] { _profile });
                    }
                    ToggleEditMode(false);
                    break;
                default:
                    // click-type selector buttons
                    if (_selectedRegion != null)
                    {
                        if (Enum.TryParse<QuickClickClickType>(key, true, out var ct))
                        {
                            _selectedRegion.ClickType = ct;
                            Invalidate();
                        }
                    }
                    break;
            }
        }
        #endregion

        #region Utilities
        private Rectangle RegionToScreenRect(QuickClickRegion r)
        {
            // Get window rect in screen coordinates
            if (_targetWindow == IntPtr.Zero)
            {
                // fallback to virtual screen
                return new Rectangle(r.X + SystemInformation.VirtualScreen.Left, r.Y + SystemInformation.VirtualScreen.Top, r.Width, r.Height);
            }

            var rect = new Win32ApiHelper.RECT();
            try
            {
                if (Win32ApiHelper.GetWindowRect(_targetWindow, ref rect))
                {
                    var winLeft = rect.Left;
                    var winTop = rect.Top;
                    return new Rectangle(winLeft + r.X, winTop + r.Y, r.Width, r.Height);
                }
            }
            catch { }

            return new Rectangle(r.X, r.Y, r.Width, r.Height);
        }

        private Point ScreenToWindowClient(Point screenPoint)
        {
            if (_targetWindow == IntPtr.Zero) return screenPoint;
            var rect = new Win32ApiHelper.RECT();
            try
            {
                if (Win32ApiHelper.GetWindowRect(_targetWindow, ref rect))
                {
                    return new Point(screenPoint.X - rect.Left, screenPoint.Y - rect.Top);
                }
            }
            catch { }
            return screenPoint;
        }

        private void RefreshPositionAndInvalidate()
        {
            if (_targetWindow != IntPtr.Zero)
            {
                var rect = new Win32ApiHelper.RECT();
                if (Win32ApiHelper.GetWindowRect(_targetWindow, ref rect))
                {
                    _targetWindowRect = new Rectangle(rect.Left, rect.Top, rect.Right - rect.Left, rect.Bottom - rect.Top);
                }
            }
            // keep overlay covering virtual screen so mapping remains consistent
            Bounds = SystemInformation.VirtualScreen;
            Invalidate();
        }

        private void ToggleEditMode(bool enable)
        {
            _editMode = enable;
            // When entering edit mode we must remove WS_EX_TRANSPARENT so the form receives mouse events
            EnsureClickThrough(!enable);
            if (enable)
            {
                // ensure focus so keyboard works
                try { this.Activate(); this.Focus(); } catch { }
            }
            Invalidate();
        }

        private void EnsureClickThrough(bool clickThrough)
        {
            // Toggle WS_EX_TRANSPARENT flag at runtime
            const int GWL_EXSTYLE = -20;
            const int WS_EX_LAYERED = 0x80000;
            const int WS_EX_TRANSPARENT = 0x20;
            const int WS_EX_NOACTIVATE = 0x08000000;
            const int WS_EX_TOOLWINDOW = 0x00000080;

            try
            {
                var ex = Win32ApiHelper.GetWindowLong(this.Handle, GWL_EXSTYLE);
                if (clickThrough)
                {
                    ex |= WS_EX_LAYERED | WS_EX_TRANSPARENT | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW;
                }
                else
                {
                    ex &= ~WS_EX_TRANSPARENT;
                    ex |= WS_EX_LAYERED | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW;
                }
                // SetWindowLong is not exposed by Win32ApiHelper; use Win32 directly
                SetWindowLong(this.Handle, GWL_EXSTYLE, ex);
            }
            catch { }
        }
        #endregion

        #region Native helpers
        [DllImport("user32.dll", SetLastError = true)]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        #endregion
    }
}
