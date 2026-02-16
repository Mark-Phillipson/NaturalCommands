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
        
        // Path to file that tracks manual hide state (persists across processes)
        private static readonly string _manualHideFlagFile = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, ".quick_clicks_manually_hidden");

        private QuickClickProfile? _profile;
        private IntPtr _targetWindow = IntPtr.Zero;
        private Rectangle _targetWindowRect = Rectangle.Empty;

        private System.Windows.Forms.Timer _repositionTimer;
        private System.Windows.Forms.Timer _hideCheckTimer;
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
            BackColor = Color.Magenta;  // Bright color for transparency key
            TransparencyKey = Color.Magenta;
            StartPosition = FormStartPosition.Manual;
            DoubleBuffered = true;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer, true);

            // Initial size - will be updated when target window is set
            Width = 800;
            Height = 600;
            Location = new Point(0, 0);

            // Timer to poll window location so overlay follows the target window
            _repositionTimer = new System.Windows.Forms.Timer { Interval = 200 };
            _repositionTimer.Tick += (s, e) => RefreshPositionAndInvalidate();

            // Timer to periodically check if another process set the manual hide flag
            _hideCheckTimer = new System.Windows.Forms.Timer { Interval = 200 };
            _hideCheckTimer.Tick += (s, e) => CheckManualHideFlag();

            // Mouse handlers for edit interactions
            MouseDown += QuickClickOverlayForm_MouseDown;
            MouseMove += QuickClickOverlayForm_MouseMove;
            MouseUp += QuickClickOverlayForm_MouseUp;
            MouseClick += QuickClickOverlayForm_MouseClick;
            KeyDown += QuickClickOverlayForm_KeyDown;

            // Make sure the overlay stays topmost
            Shown += (s, e) => {
                try
                {
                    Win32ApiHelper.SetWindowPos(Handle, new IntPtr(-1), 0, 0, 0, 0, 0x0001 | 0x0002);
                }
                catch { }
            };
        }
        
        #region Public API (thread-safe)
        public static void InitializeUIContext()
        {
            _uiContext = System.Threading.SynchronizationContext.Current;
            Logger.LogDebug($"QuickClickOverlayForm: UI context initialized: {_uiContext?.GetType().Name ?? "null"}");
        }

        public static void ShowOverlay(QuickClickProfile? profile, IntPtr windowHandle, bool enterEditMode = false)
        {
            if (_uiContext != null && System.Threading.SynchronizationContext.Current != _uiContext)
            {
                _uiContext.Post(_ => ShowOverlay(profile, windowHandle, enterEditMode), null);
                return;
            }

            lock (_lock)
            {
                SetManuallyHidden(false); // Clear manual hide flag when showing
                Logger.LogDebug($"QuickClickOverlayForm.ShowOverlay: Manual hide flag cleared, showing overlay (EditMode={enterEditMode})");
                if (_instance == null || _instance.IsDisposed)
                {
                    _instance = new QuickClickOverlayForm();
                }

                _instance.InternalShow(profile, windowHandle, enterEditMode);
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
                SetManuallyHidden(true); // Set manual hide flag (persists across processes via file)
                Logger.LogDebug($"QuickClickOverlayForm.HideOverlay: Manual hide flag set");
                if (_instance != null && !_instance.IsDisposed)
                {
                    Logger.LogDebug($"QuickClickOverlayForm.HideOverlay: Instance exists, calling InternalHide (Visible={_instance.Visible})");
                    _instance.InternalHide();
                    Logger.LogDebug($"QuickClickOverlayForm.HideOverlay: Overlay hidden (Visible={_instance.Visible})");
                }
                else
                {
                    Logger.LogDebug($"QuickClickOverlayForm.HideOverlay: No instance to hide (_instance null={_instance == null}, disposed={_instance?.IsDisposed})");
                }
            }
        }

        public static bool IsVisible => _instance != null && !_instance.IsDisposed && _instance.Visible;
        
        public static bool IsInEditMode => _instance != null && !_instance.IsDisposed && _instance._editMode;
        
        public static bool IsManuallyHidden 
        {
            get
            {
                try { return System.IO.File.Exists(_manualHideFlagFile); }
                catch { return false; }
            }
        }
        
        private static void SetManuallyHidden(bool hidden)
        {
            try
            {
                if (hidden)
                {
                    System.IO.File.WriteAllText(_manualHideFlagFile, DateTime.Now.ToString());
                    Logger.LogDebug($"QuickClickOverlayForm: Manual hide flag file CREATED");
                }
                else
                {
                    if (System.IO.File.Exists(_manualHideFlagFile))
                    {
                        System.IO.File.Delete(_manualHideFlagFile);
                        Logger.LogDebug($"QuickClickOverlayForm: Manual hide flag file DELETED");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"QuickClickOverlayForm: Error setting manual hide flag: {ex.Message}");
            }
        }

        public static void ClearManualHideFlag()
        {
            SetManuallyHidden(false);
        }

        public static void EnterEditMode()
        {
            if (_uiContext != null && System.Threading.SynchronizationContext.Current != _uiContext)
            {
                _uiContext.Post(_ => EnterEditMode(), null);
                return;
            }

            lock (_lock)
            {
                SetManuallyHidden(false); // Clear manual hide flag when entering edit mode
                if (_instance != null && !_instance.IsDisposed && _instance.Visible)
                {
                    _instance.ToggleEditMode(true);
                }
                else
                {
                    Logger.LogWarning($"QuickClickOverlayForm.EnterEditMode: Cannot enter edit mode - instance is null, disposed, or not visible");
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
                SetManuallyHidden(false); // Clear manual hide flag when creating new region
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
        private void InternalShow(QuickClickProfile? profile, IntPtr windowHandle, bool enterEditMode = false)
        {
            // Don't show if manually hidden (defensive check)
            if (IsManuallyHidden)
            {
                Logger.LogDebug($"QuickClickOverlayForm.InternalShow: Skipping show because manually hidden");
                return;
            }

            _profile = profile;
            _targetWindow = windowHandle;
            RefreshPositionAndInvalidate();

            // Set edit mode BEFORE showing to avoid window recreation flicker
            ToggleEditMode(enterEditMode);
            _selectedRegion = null;
            _placementMode = false;
            _isDragging = false;

            _repositionTimer.Start();
            _hideCheckTimer.Start();

            if (!Visible)
                Show();
            BringToFront();
            Logger.LogDebug($"QuickClickOverlayForm.InternalShow: Overlay shown (Visible={Visible}, EditMode={_editMode})");
        }

        private void InternalHide()
        {
            _repositionTimer.Stop();
            _hideCheckTimer.Stop();
            _placementMode = false;
            _isDragging = false;
            _selectedRegion = null;
            ToggleEditMode(false);
            
            // Force hide with both methods to ensure it actually hides
            if (Visible)
            {
                Logger.LogDebug($"QuickClickOverlayForm.InternalHide: Hiding visible overlay");
                Visible = false;
                Hide();
            }
            else
            {
                Logger.LogDebug($"QuickClickOverlayForm.InternalHide: Overlay already not visible");
            }
        }
        #endregion

        #region Painting
        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            var g = e.Graphics;
            
            // Don't clear - transparent background (black) will show through
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            // Draw warning if no profile provided
            if (_profile == null)
            {
                var msg = "No Quick Click profile for this app/monitor. Click '+ New' to create a region.";
                var size = g.MeasureString(msg, _labelFont);
                var rect = new RectangleF(20, 60, size.Width + 16, size.Height + 10);
                using (var bg = new SolidBrush(Color.FromArgb(220, 40, 40, 120))) g.FillRectangle(bg, rect);
                g.DrawString(msg, _labelFont, Brushes.White, rect.X + 8, rect.Y + 4);
                if (!_editMode) return;
            }

            // Draw regions
            if (_profile != null)
            {
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
            // Don't fill the background - let magenta (transparency key) show through

            // Title banner at the very top - make it VERY visible
            using (var bannerBrush = new SolidBrush(Color.FromArgb(255, 0, 100, 200)))
            {
                var bannerRect = new Rectangle(0, 0, Width, 80);
                g.FillRectangle(bannerBrush, bannerRect);
                
                // Draw a bright border around the banner
                using (var borderPen = new Pen(Color.Yellow, 4))
                {
                    g.DrawRectangle(borderPen, 2, 2, Width - 4, 76);
                }
                
                var titleFont = new Font("Segoe UI", 16, FontStyle.Bold);
                var title = "QUICK CLICKS EDITOR - EDIT MODE ACTIVE";
                g.DrawString(title, titleFont, Brushes.White, 30, 20);
                
                var subtitle = $"Screen: {Width}x{Height} - Say: 'save and exit', 'new', 'delete'";
                g.DrawString(subtitle, _labelFont, Brushes.Yellow, 30, 50);
                titleFont.Dispose();
            }

            // Toolbar below the banner - make buttons MUCH larger and more visible
            int x = 30; int y = 100; int h = 50; int gap = 15;
            _editorButtons.Clear();

            void DrawBtn(string label, Color bg)
            {
                var btnFont = new Font("Segoe UI", 14, FontStyle.Bold);
                var size = g.MeasureString(label, btnFont);
                var rect = new Rectangle(x, y, (int)size.Width + 40, h);
                
                // Draw button background
                using (var b = new SolidBrush(bg)) g.FillRectangle(b, rect);
                
                // Draw bright white border
                using (var border = new Pen(Color.White, 3)) g.DrawRectangle(border, rect);
                
                // Draw text
                g.DrawString(label, btnFont, Brushes.White, rect.X + 20, rect.Y + 12);
                _editorButtons.Add((label, rect));
                x += rect.Width + gap;
                btnFont.Dispose();
            }

            DrawBtn("+ New", Color.FromArgb(255, 120, 50, 200));
            DrawBtn("+ Rectangle", Color.FromArgb(255, 50, 180, 50));
            DrawBtn("Delete", Color.FromArgb(255, 220, 30, 30));
            DrawBtn("Save & Exit", Color.FromArgb(255, 30, 150, 220));

            // If a region is selected, draw click-type selector below buttons
            if (_selectedRegion != null)
            {
                var labelText = $"SELECTED REGION: {_selectedRegion.Name} - Choose Click Type:";
                var labelFont = new Font("Segoe UI", 12, FontStyle.Bold);
                g.DrawString(labelText, labelFont, Brushes.Yellow, 30, y + h + 20);
                labelFont.Dispose();
                
                var badgeX = 30; var badgeY = y + h + 55;
                var types = new[] { "Left", "Double", "Right" };
                foreach (var t in types)
                {
                    var rect = new Rectangle(badgeX, badgeY, 120, 45);
                    var isActive = string.Equals(_selectedRegion.ClickType.ToString(), t, StringComparison.OrdinalIgnoreCase);
                    using (var b = new SolidBrush(isActive ? Color.FromArgb(255, 0, 200, 100) : Color.FromArgb(255, 80, 80, 80))) 
                        g.FillRectangle(b, rect);
                    using (var border = new Pen(isActive ? Color.Yellow : Color.White, isActive ? 4 : 2)) 
                        g.DrawRectangle(border, rect);
                    var font = new Font("Segoe UI", 12, isActive ? FontStyle.Bold : FontStyle.Regular);
                    g.DrawString(t, font, Brushes.White, rect.X + 20, rect.Y + 12);
                    font.Dispose();
                    _editorButtons.Add((t, rect));
                    badgeX += rect.Width + 15;
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
            // Since our overlay is positioned exactly on the target window,
            // regions are just at their X, Y offsets from the form's top-left
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
                    // Position overlay exactly on the target window
                    Bounds = _targetWindowRect;
                }
            }
            Invalidate();
        }

        private void CheckManualHideFlag()
        {
            // Check if another process has set the manual hide flag
            if (IsManuallyHidden && Visible)
            {
                Logger.LogDebug("QuickClickOverlayForm.CheckManualHideFlag: Manual hide flag detected from another process - hiding overlay");
                InternalHide();
            }
        }

        private void ToggleEditMode(bool enable)
        {
            if (_editMode == enable) return; // Already in desired mode
            
            _editMode = enable;
            // Force window recreation to apply new CreateParams
            if (IsHandleCreated)
            {
                RecreateHandle();
            }
            if (enable)
            {
                // ensure focus so keyboard works
                try { this.Activate(); this.Focus(); } catch { }
            }
            Invalidate();
        }

        private void EnsureClickThrough(bool clickThrough)
        {
            // This method is no longer needed - CreateParams handles it at window creation
            // Keeping it as a no-op for compatibility
        }
        #endregion
    }
}
