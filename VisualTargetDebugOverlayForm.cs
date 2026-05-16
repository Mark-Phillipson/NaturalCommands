using System;
using System.Drawing;
using System.Windows.Forms;

namespace NaturalCommands
{
    /// <summary>
    /// Debug overlay that displays OCR/UI automation screenshot and results
    /// when visual targeting finds zero candidates.
    /// </summary>
    public class VisualTargetDebugOverlayForm : Form
    {
        private static VisualTargetDebugOverlayForm? _instance;
        private Label labelQuery = null!;
        private Label labelResult = null!;
        private PictureBox pictureBox = null!;

        public VisualTargetDebugOverlayForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.BackColor = Color.Black;
            this.ForeColor = Color.White;
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            this.TopMost = true;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.Manual;
            this.ControlBox = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Title bar
            labelQuery = new Label
            {
                Text = "Visual Target Debug: [initializing]",
                ForeColor = Color.Cyan,
                BackColor = Color.Black,
                Font = new Font("Consolas", 11, FontStyle.Bold),
                AutoSize = false,
                Dock = DockStyle.Top,
                Height = 30,
                Padding = new Padding(8)
            };
            this.Controls.Add(labelQuery);

            // Screenshot display
            pictureBox = new PictureBox
            {
                BackColor = Color.Gray,
                SizeMode = PictureBoxSizeMode.Zoom,
                Dock = DockStyle.Fill
            };
            this.Controls.Add(pictureBox);

            // Result message
            labelResult = new Label
            {
                Text = "[no results]",
                ForeColor = Color.Yellow,
                BackColor = Color.Black,
                Font = new Font("Consolas", 10),
                AutoSize = false,
                Dock = DockStyle.Bottom,
                Height = 60,
                Padding = new Padding(8),
                BorderStyle = BorderStyle.Fixed3D
            };
            this.Controls.Add(labelResult);
        }

        /// <summary>
        /// Show debug overlay with screenshot and search results.
        /// </summary>
        public static void ShowDebug(string query, Image? screenshot, string resultMessage, int timeoutMs = 5000)
        {
            try
            {
                NaturalCommands.Helpers.Logger.LogInfo($"ShowDebug called for query: '{query}', screenshot: {screenshot != null}, timeout: {timeoutMs}ms");
                HideDebug();

                _instance = new VisualTargetDebugOverlayForm();
                _instance.labelQuery.Text = $"Visual Target Debug: {query}";
                _instance.labelResult.Text = resultMessage;

                if (screenshot != null)
                {
                    _instance.pictureBox.Image = screenshot;
                    NaturalCommands.Helpers.Logger.LogInfo($"Screenshot set: {screenshot.Width}x{screenshot.Height}");
                }

                // Position at top-right, 50% of primary screen size
                var screen = Screen.PrimaryScreen;
                if (screen != null)
                {
                    int width = Math.Max(400, screen.Bounds.Width / 2);
                    int height = Math.Max(300, screen.Bounds.Height / 2);
                    _instance.Width = width;
                    _instance.Height = height;
                    _instance.Left = screen.Bounds.Right - width - 20;
                    _instance.Top = screen.Bounds.Top + 80;
                    NaturalCommands.Helpers.Logger.LogInfo($"Form positioned: {width}x{height} at ({_instance.Left}, {_instance.Top})");
                }

                _instance.Show();
                NaturalCommands.Helpers.Logger.LogInfo($"ShowDebug form displayed, will auto-close in {timeoutMs}ms");

                // Auto-close after timeout
                var timer = new System.Windows.Forms.Timer { Interval = timeoutMs };
                timer.Tick += (s, e) =>
                {
                    timer.Stop();
                    timer.Dispose();
                    NaturalCommands.Helpers.Logger.LogInfo("Debug overlay auto-closing");
                    HideDebug();
                };
                timer.Start();
            }
            catch (Exception ex)
            {
                NaturalCommands.Helpers.Logger.LogError($"VisualTargetDebugOverlayForm.ShowDebug failed: {ex.Message}\n{ex.StackTrace}");
            }
        }

        /// <summary>
        /// Hide and dispose debug overlay.
        /// </summary>
        public static void HideDebug()
        {
            try
            {
                if (_instance != null)
                {
                    _instance.Hide();
                    _instance.Dispose();
                    _instance = null;
                }
            }
            catch { }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (pictureBox?.Image != null)
                {
                    pictureBox.Image.Dispose();
                }
            }
            base.Dispose(disposing);
        }
    }
}
