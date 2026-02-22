using System;
using System.Collections.Generic;
using System.Drawing;
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

        public static void InitializeUIContext()
        {
            _uiContext = System.Threading.SynchronizationContext.Current;
        }

        public static bool IsVisible => _instance != null && !_instance.IsDisposed && _instance.Visible;

        public static void ShowCandidates(List<VisualTargetCandidate> candidates, int timeoutMs)
        {
            if (_uiContext != null && System.Threading.SynchronizationContext.Current != _uiContext)
            {
                _uiContext.Post(_ => ShowCandidates(candidates, timeoutMs), null);
                return;
            }

            lock (_lock)
            {
                if (_instance == null || _instance.IsDisposed)
                {
                    _instance = new VisualCandidateOverlayForm();
                }

                _instance._candidates = candidates ?? new List<VisualTargetCandidate>();
                _instance._timeoutTimer.Stop();
                _instance._timeoutTimer.Interval = Math.Max(1000, timeoutMs);
                _instance._timeoutTimer.Start();

                if (!_instance.Visible)
                {
                    _instance.Show();
                }

                _instance.BringToFront();
                _instance.Invalidate();
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
                var center = candidate.Center;
                var centerInOverlay = new Point(center.X - virtualX, center.Y - virtualY);

                var markerRect = new Rectangle(centerInOverlay.X - 18, centerInOverlay.Y - 18, 36, 36);
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
