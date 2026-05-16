using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using NaturalCommands.Helpers;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Windows.Forms;
using NaturalCommands.Models;

namespace NaturalCommands
{
    public class VisualCandidateOverlayForm : Form
    {
        private const int WM_NCHITTEST = 0x84;
        private const int HTTRANSPARENT = -1;
        private const int WS_EX_NOACTIVATE = 0x08000000;
        private const int WS_EX_TRANSPARENT = 0x00000020;
        private const int WS_EX_TOOLWINDOW = 0x00000080;

        private static VisualCandidateOverlayForm? _instance;
        private static readonly object _lock = new();
        private static System.Threading.SynchronizationContext? _uiContext;

        private List<VisualTargetCandidate> _candidates = new();
        private readonly Font _numberFont = new("Segoe UI", 12, FontStyle.Bold);
        private readonly Font _hintFont = new("Segoe UI", 10, FontStyle.Regular);
        private readonly System.Windows.Forms.Timer _timeoutTimer;

        private VisualCandidateOverlayForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            BackColor = Color.Lime;
            TransparencyKey = Color.Lime;
            StartPosition = FormStartPosition.Manual;
            KeyPreview = true;
            DoubleBuffered = true;

            var bounds = Screen.AllScreens.Aggregate(Rectangle.Empty, (current, screen) => Rectangle.Union(current, screen.Bounds));
            Bounds = bounds;

            _timeoutTimer = new System.Windows.Forms.Timer();
            _timeoutTimer.Interval = 9000;
            _timeoutTimer.Tick += (s, e) => HideOverlay();

            KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Escape)
                {
                    HideOverlay();
                }
            };
        }

        // Low-level keyboard hook to capture numeric key presses while overlay is visible
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_SYSKEYDOWN = 0x0104;

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        private static LowLevelKeyboardProc? _keyboardProc;
        private static IntPtr _hookId = IntPtr.Zero;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, int dwExtraInfo);

        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;

        public static void InitializeUIContext()
        {
            _uiContext = System.Threading.SynchronizationContext.Current;
        }

        public static bool IsVisible => _instance != null && !_instance.IsDisposed && _instance.Visible;

        public static void ShowCandidates(List<VisualTargetCandidate> candidates, int timeoutMs)
        {
            try
            {
                NaturalCommands.Helpers.Logger.LogInfo($"ShowCandidates called with {candidates?.Count ?? 0} candidates, timeout={timeoutMs}ms");
                
                if (_uiContext != null && System.Threading.SynchronizationContext.Current != _uiContext)
                {
                    NaturalCommands.Helpers.Logger.LogInfo("ShowCandidates: posting to UI context");
                    _uiContext.Post(_ => ShowCandidates(candidates, timeoutMs), null);
                    return;
                }

                lock (_lock)
                {
                    if (_instance == null || _instance.IsDisposed)
                    {
                        _instance = new VisualCandidateOverlayForm();
                        NaturalCommands.Helpers.Logger.LogInfo("ShowCandidates: created new form instance");
                    }

                    _instance._candidates = candidates ?? new List<VisualTargetCandidate>();
                    _instance._timeoutTimer.Stop();
                    _instance._timeoutTimer.Interval = Math.Max(1000, timeoutMs);
                    _instance._timeoutTimer.Start();

                    if (!_instance.Visible)
                    {
                        _instance.Show();
                        NaturalCommands.Helpers.Logger.LogInfo("ShowCandidates: form shown");
                    }
                    EnsureKeyboardHookInstalled();
                    
                    _instance.BringToFront();
                    _instance.Invalidate();
                    NaturalCommands.Helpers.Logger.LogInfo("ShowCandidates: form brought to front and invalidated");
                }
            }
            catch (Exception ex)
            {
                NaturalCommands.Helpers.Logger.LogError($"ShowCandidates exception: {ex.Message}\n{ex.StackTrace}");
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
                if (_instance == null || _instance.IsDisposed)
                {
                    return;
                }

                _instance._timeoutTimer.Stop();
                _instance.Hide();
                EnsureKeyboardHookRemoved();
            }
        }

        private static void EnsureKeyboardHookInstalled()
        {
            try
            {
                if (_hookId != IntPtr.Zero) return;
                _keyboardProc = KeyboardHookCallback;
                using var curProcess = System.Diagnostics.Process.GetCurrentProcess();
                using var curModule = curProcess.MainModule;
                var moduleHandle = GetModuleHandle(curModule?.ModuleName ?? string.Empty);
                _hookId = SetWindowsHookEx(WH_KEYBOARD_LL, _keyboardProc!, moduleHandle, 0);
                Logger.LogDebug($"VisualCandidateOverlayForm: keyboard hook installed (hookId={_hookId})");
            }
            catch (Exception ex)
            {
                Logger.LogError($"VisualCandidateOverlayForm: failed to install keyboard hook: {ex.Message}");
            }
        }

        private static void EnsureKeyboardHookRemoved()
        {
            try
            {
                if (_hookId == IntPtr.Zero) return;
                UnhookWindowsHookEx(_hookId);
                _hookId = IntPtr.Zero;
                _keyboardProc = null;
                Logger.LogDebug("VisualCandidateOverlayForm: keyboard hook removed");
            }
            catch (Exception ex)
            {
                Logger.LogError($"VisualCandidateOverlayForm: failed to remove keyboard hook: {ex.Message}");
            }
        }

        private static IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (nCode >= 0 && (wParam == (IntPtr)WM_KEYDOWN || wParam == (IntPtr)WM_SYSKEYDOWN))
                {
                    int vkCode = Marshal.ReadInt32(lParam);
                    int num = -1;
                    // Top-row digits '0'..'9'
                    if (vkCode >= 0x30 && vkCode <= 0x39)
                        num = vkCode - 0x30;
                    // Numpad digits
                    else if (vkCode >= 0x60 && vkCode <= 0x69)
                        num = vkCode - 0x60;

                    // ESC -> cancel overlay
                    if (vkCode == 0x1B)
                    {
                        _instance?.BeginInvoke(new Action(() => HideOverlay()));
                        return (IntPtr)1;
                    }

                    if (num >= 1)
                    {
                        // Dispatch to instance to handle safe UI interaction
                        _instance?.BeginInvoke(new Action(() => _instance?.HandleNumberKey(num)));
                        // Swallow the key so underlying app doesn't receive it
                        return (IntPtr)1;
                    }
                }
            }
            catch { }
            return CallNextHookEx(_hookId, nCode, wParam, lParam);
        }

        private void HandleNumberKey(int number)
        {
            try
            {
                if (!Helpers.VisualCandidateSessionStore.TryGetCandidate(number, out var candidate) || candidate == null)
                {
                    // no candidate for this number
                    Logger.LogDebug($"No visual candidate for number {number}");
                    return;
                }

                // Close/hide overlay first so the click reaches underlying window
                HideOverlay();
                System.Threading.Thread.Sleep(30);

                var session = Helpers.VisualCandidateSessionStore.GetSession();
                var point = Helpers.VisualTargetClickPointResolver.Resolve(candidate, session?.Query ?? string.Empty);
                if (point == System.Drawing.Point.Empty)
                    point = candidate.Center;

                Logger.LogInfo($"Visual overlay: keyed choose {number} -> click point ({point.X},{point.Y}) for '{candidate.Label}'");

                System.Drawing.Point prev = System.Windows.Forms.Cursor.Position;
                SetCursorPos(point.X, point.Y);
                System.Threading.Thread.Sleep(35);
                mouse_event(MOUSEEVENTF_LEFTDOWN, point.X, point.Y, 0, 0);
                System.Threading.Thread.Sleep(20);
                mouse_event(MOUSEEVENTF_LEFTUP, point.X, point.Y, 0, 0);
                System.Threading.Thread.Sleep(70);

                try { SetCursorPos(prev.X, prev.Y); } catch { }

                Helpers.VisualCandidateSessionStore.Clear();
            }
            catch (Exception ex)
            {
                Logger.LogError($"HandleNumberKey error: {ex.Message}");
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            var graphics = e.Graphics;
            graphics.SmoothingMode = SmoothingMode.AntiAlias;

            int virtualX = SystemInformation.VirtualScreen.Left;
            int virtualY = SystemInformation.VirtualScreen.Top;

            for (int i = 0; i < _candidates.Count; i++)
            {
                var candidate = _candidates[i];
                // Position marker near the candidate bounding rect top-right (more visually stable
                // for grid-like tiles). Compute candidate rect relative to the overlay virtual origin.
                var candidateRectInOverlay = new Rectangle(
                    candidate.Bounds.Left - virtualX,
                    candidate.Bounds.Top - virtualY,
                    candidate.Bounds.Width,
                    candidate.Bounds.Height);

                int markerLeft = Math.Max(2, candidateRectInOverlay.Right - 28);
                int markerTop = Math.Max(2, candidateRectInOverlay.Top + 6);
                var markerRect = new Rectangle(markerLeft, markerTop, 36, 36);

                // Draw candidate bounding rect for clarity
                Color rectColor = string.Equals(candidate.Source, "uia", StringComparison.OrdinalIgnoreCase)
                    ? Color.FromArgb(200, 0, 200, 0)
                    : Color.FromArgb(200, 34, 139, 230);
                using var rectPen = new Pen(rectColor, 2) { DashStyle = DashStyle.Solid };
                var outlineRect = candidateRectInOverlay;
                // inflate slightly so the rectangle is visible
                outlineRect.Inflate(2, 2);
                graphics.DrawRectangle(rectPen, outlineRect);

                // Draw label snippet near the marker
                var label = candidate.Label ?? string.Empty;
                var shortLabel = label.Length > 28 ? label.Substring(0, 25) + "..." : label;
                var labelSize = graphics.MeasureString(shortLabel, _hintFont);
                var labelBg = new RectangleF(markerRect.Right + 6, markerRect.Y - 2, labelSize.Width + 8, labelSize.Height + 4);
                using var labelFill = new SolidBrush(Color.FromArgb(210, 0, 0, 0));
                graphics.FillRectangle(labelFill, labelBg);
                graphics.DrawString(shortLabel, _hintFont, Brushes.White, labelBg.X + 4, labelBg.Y + 2);

                using var fill = new SolidBrush(Color.FromArgb(220, 34, 139, 230));
                using var border = new Pen(Color.FromArgb(240, 255, 255, 255), 2);
                graphics.FillEllipse(fill, markerRect);
                graphics.DrawEllipse(border, markerRect);

                string text = (i + 1).ToString();
                var textSize = graphics.MeasureString(text, _numberFont);
                graphics.DrawString(text, _numberFont, Brushes.White,
                    markerRect.X + (markerRect.Width - textSize.Width) / 2,
                    markerRect.Y + (markerRect.Height - textSize.Height) / 2 - 1);
            }

            string instruction = "Say 'choose <number>' or press ESC to cancel";
            var instructionSize = graphics.MeasureString(instruction, _hintFont);
            var rect = new RectangleF((Width - instructionSize.Width) / 2 - 8, Height - 56, instructionSize.Width + 16, instructionSize.Height + 10);
            using var bg = new SolidBrush(Color.FromArgb(190, 0, 0, 0));
            graphics.FillRectangle(bg, rect);
            graphics.DrawString(instruction, _hintFont, Brushes.White, rect.X + 8, rect.Y + 5);
        }

        protected override bool ShowWithoutActivation => true;

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= WS_EX_NOACTIVATE;
                cp.ExStyle |= WS_EX_TRANSPARENT;
                cp.ExStyle |= WS_EX_TOOLWINDOW;
                return cp;
            }
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_NCHITTEST)
            {
                m.Result = (IntPtr)HTTRANSPARENT;
                return;
            }

            base.WndProc(ref m);
        }
    }
}
