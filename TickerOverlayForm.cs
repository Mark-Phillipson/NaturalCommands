using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace NaturalCommands
{
    public enum TickerCategory
    {
        Info,
        Warning,
        Critical,
        Success
    }

    public sealed record TickerMessage(string Text, TickerCategory Category = TickerCategory.Info);

    public partial class TickerOverlayForm : Form
    {
        private readonly List<TickerMessage> _messages;
        private int _currentIndex;
        private int _cycleCount;
        private readonly int _maxCycles;
        private readonly bool _hideOnDismiss;
        private readonly System.Windows.Forms.Timer _cycleTimer;
        private readonly System.Windows.Forms.Timer _criticalPulseTimer;
        private bool _pulseVisible;

        public TickerOverlayForm(IEnumerable<string> lines, int cycleSeconds = 5, int maxCycles = 3, bool topPosition = false, bool hideOnDismiss = false)
            : this(ParseMessages(lines), cycleSeconds, maxCycles, topPosition, hideOnDismiss)
        {
        }

        public TickerOverlayForm(IEnumerable<TickerMessage> messages, int cycleSeconds = 5, int maxCycles = 3, bool topPosition = false, bool hideOnDismiss = false)
        {
            _messages = messages?.Where(m => !string.IsNullOrWhiteSpace(m.Text)).ToList() ?? new List<TickerMessage>();
            if (_messages.Count == 0)
            {
                _messages.Add(new TickerMessage("No notifications", TickerCategory.Info));
            }

            _maxCycles = maxCycles;  // 0 means never auto-close; >0 means auto-close after N cycles
            _hideOnDismiss = hideOnDismiss;
            InitializeComponent();

            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            ShowInTaskbar = false;
            TopMost = true;
            BackColor = Color.FromArgb(30, 30, 30);

            var screen = Screen.PrimaryScreen?.WorkingArea ?? Screen.AllScreens[0].WorkingArea;
            Width = screen.Width;
            Height = 140;
            Left = screen.Left;
            Top = topPosition ? screen.Top : screen.Bottom - Height;

            _cycleTimer = new System.Windows.Forms.Timer { Interval = cycleSeconds * 1000 };
            _cycleTimer.Tick += CycleTimer_Tick;
            _criticalPulseTimer = new System.Windows.Forms.Timer { Interval = 500 };
            _criticalPulseTimer.Tick += CriticalPulseTimer_Tick;

            _labelMessage.Click += (s, e) => NextMessage();
            _buttonNext.Click += (s, e) => NextMessage();
            _buttonPrev.Click += (s, e) => PreviousMessage();
            _buttonDismiss.Click += (s, e) => Dismiss();

            KeyPreview = true;
            KeyDown += TickerOverlayForm_KeyDown;
            MouseEnter += (s, e) => _cycleTimer.Stop();
            MouseLeave += (s, e) => StartCycleTimerIfNotCritical();
            Load += (s, e) => ForceTopmost();

            ShowMessage(0);
            _cycleTimer.Start();
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOACTIVATE = 0x0010;

        private void ForceTopmost()
        {
            SetWindowPos(Handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
        }

        public static TickerMessage ParseSingleLine(string line)
        {
            var list = ParseMessages(new[] { line });
            return list.Count > 0 ? list[0] : new TickerMessage(line, TickerCategory.Info);
        }

        private static List<TickerMessage> ParseMessages(IEnumerable<string> lines)
        {
            var list = new List<TickerMessage>();
            if (lines == null)
            {
                return list;
            }

            foreach (var raw in lines)
            {
                if (string.IsNullOrWhiteSpace(raw))
                    continue;

                var line = raw.Trim();
                var category = TickerCategory.Info;
                var text = line;

                var colonIndex = line.IndexOf(':');
                if (colonIndex > 0)
                {
                    var prefix = line.Substring(0, colonIndex).Trim();
                    var remainder = line.Substring(colonIndex + 1).Trim();
                    if (Enum.TryParse<TickerCategory>(prefix, true, out var parsed))
                    {
                        category = parsed;
                        text = remainder;
                    }
                }

                if (string.IsNullOrWhiteSpace(text))
                    continue;

                list.Add(new TickerMessage(text, category));
            }

            if (!list.Any())
            {
                list.Add(new TickerMessage("No notifications", TickerCategory.Info));
            }

            return list;
        }

        private void StartCriticalPulse()
        {
            _criticalPulseTimer.Start();
            _cycleTimer.Stop();
        }

        private void StopCriticalPulse()
        {
            _criticalPulseTimer.Stop();
            _pulseVisible = false;
            Invalidate();
            StartCycleTimerIfNotCritical();
        }

        private void CriticalPulseTimer_Tick(object? sender, EventArgs e)
        {
            _pulseVisible = !_pulseVisible;
            Invalidate();
        }

        private void CycleTimer_Tick(object? sender, EventArgs e)
        {
            if (CurrentCategory == TickerCategory.Critical)
            {
                return;
            }

            NextMessage();
        }

        private void StartCycleTimerIfNotCritical()
        {
            if (CurrentCategory != TickerCategory.Critical && !_cycleTimer.Enabled)
            {
                _cycleTimer.Start();
            }
        }

        private void TickerOverlayForm_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                Dismiss();
            }
            else if (e.KeyCode == Keys.Right)
            {
                NextMessage();
            }
            else if (e.KeyCode == Keys.Left)
            {
                PreviousMessage();
            }
        }

        private TickerMessage CurrentMessage => _messages[_currentIndex];

        private TickerCategory CurrentCategory => CurrentMessage.Category;

        private void ShowMessage(int newIndex)
        {
            _currentIndex = (newIndex + _messages.Count) % _messages.Count;
            _cycleCount++;

            // maxCycles == 0 means never auto-close — user must dismiss explicitly
            if (_maxCycles > 0 && _cycleCount > _maxCycles * _messages.Count)
            {
                Dismiss();
                return;
            }

            var message = CurrentMessage;
            var icon = message.Category switch
            {
                TickerCategory.Warning => "⚠️",
                TickerCategory.Critical => "🔴",
                TickerCategory.Success => "✅",
                _ => "ℹ️",
            };

            _labelMessage.Text = $"{icon} [{_currentIndex + 1}/{_messages.Count}] {message.Text}";
            _labelCounter.Text = message.Category.ToString().ToUpper();
            _accentBar.BackColor = message.Category switch
            {
                TickerCategory.Warning => Color.FromArgb(255, 255, 193, 7),
                TickerCategory.Critical => Color.FromArgb(255, 244, 67, 54),
                TickerCategory.Success => Color.FromArgb(255, 76, 175, 80),
                _ => Color.FromArgb(255, 33, 150, 243),
            };

            if (message.Category == TickerCategory.Critical)
            {
                StartCriticalPulse();
            }
            else
            {
                StopCriticalPulse();
            }

            Invalidate();
        }

        private void NextMessage() => ShowMessage(_currentIndex + 1);

        private void PreviousMessage() => ShowMessage(_currentIndex - 1);

        private void Dismiss()
        {
            try { NaturalCommands.Helpers.Logger.LogInfo($"[TICKER-FORM] Dismiss invoked - hideOnDismiss={_hideOnDismiss}, Visible={Visible}"); } catch { }
            if (_hideOnDismiss)
            {
                try { NaturalCommands.Helpers.Logger.LogInfo("[TICKER-FORM] Hiding ticker (hideOnDismiss=true)"); } catch { }
                Hide();
                return;
            }

            try { NaturalCommands.Helpers.Logger.LogInfo("[TICKER-FORM] Closing ticker (hideOnDismiss=false)"); } catch { }
            Close();
        }

        private void EnsureVisible()
        {
            if (!Visible)
            {
                try { NaturalCommands.Helpers.Logger.LogInfo("[TICKER-FORM] EnsureVisible: showing ticker (was hidden)"); } catch { }
                try { Show(); } catch { }
            }

            if (WindowState == FormWindowState.Minimized)
            {
                WindowState = FormWindowState.Normal;
            }

            try
            {
                // Make sure the window is topmost and in front. Try to activate if possible.
                TopMost = true;
                BringToFront();
                ForceTopmost();
                try { Activate(); } catch { }
            }
            catch { }
        }

        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);
            try { NaturalCommands.Helpers.Logger.LogInfo($"[TICKER-FORM] Visible changed -> {Visible}"); } catch { }
        }

        /// <summary>
        /// Thread-safe: adds a new message to the running ticker and refreshes the display.
        /// Safe to call from any thread.
        /// </summary>
        public void AddMessage(TickerMessage message)
        {
            if (InvokeRequired)
            {
                try { BeginInvoke(new Action(() => AddMessage(message))); } catch { }
                return;
            }

            _messages.Add(message);
            try { NaturalCommands.Helpers.Logger.LogInfo($"[TICKER-FORM] AddMessage: queued message '{message.Text}' (Category={message.Category})"); } catch { }
            // Reset cycle counter so the new message gets a full set of cycles
            _cycleCount = 0;
            EnsureVisible();
            // Show the newly added message immediately
            ShowMessage(_messages.Count - 1);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            try { NaturalCommands.Helpers.Logger.LogInfo($"[TICKER-FORM] OnFormClosing: reason={e.CloseReason}, hideOnDismiss={_hideOnDismiss}"); } catch { }
            if (_hideOnDismiss && e.CloseReason == CloseReason.UserClosing)
            {
                try { NaturalCommands.Helpers.Logger.LogInfo("[TICKER-FORM] OnFormClosing: canceling close and hiding instead"); } catch { }
                e.Cancel = true;
                Hide();
                return;
            }

            base.OnFormClosing(e);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            using var brush = new SolidBrush(BackColor);
            e.Graphics.FillRectangle(brush, ClientRectangle);

            using var accentBrush = new SolidBrush(_accentBar.BackColor);
            e.Graphics.FillRectangle(accentBrush, 0, 0, 8, Height);

            if (CurrentCategory == TickerCategory.Critical && _pulseVisible)
            {
                using var pen = new Pen(Color.Red, 6);
                e.Graphics.DrawRectangle(pen, 2, 2, Width - 5, Height - 5);
            }
        }
    }
}
