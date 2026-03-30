using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace NaturalCommands
{
    public partial class TickerOverlayForm
    {
        private IContainer components;
        private Label _labelMessage;
        private Label _labelCounter;
        private Button _buttonPrev;
        private Button _buttonNext;
        private Button _buttonDismiss;
        private Panel _accentBar;

        private void InitializeComponent()
        {
            components = new Container();
            _labelMessage = new Label();
            _labelCounter = new Label();
            _buttonPrev = new Button();
            _buttonNext = new Button();
            _buttonDismiss = new Button();
            _accentBar = new Panel();

            SuspendLayout();

            // _accentBar
            _accentBar.BackColor = Color.SteelBlue;
            _accentBar.Location = new Point(0, 0);
            _accentBar.Size = new Size(8, 80);

            // _labelMessage
            _labelMessage.AutoSize = false;
            _labelMessage.Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point);
            _labelMessage.ForeColor = Color.White;
            _labelMessage.Location = new Point(14, 12);
            _labelMessage.Size = new Size(920, 48);
            _labelMessage.TextAlign = ContentAlignment.MiddleLeft;

            // _labelCounter
            _labelCounter.AutoSize = true;
            _labelCounter.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            _labelCounter.ForeColor = Color.WhiteSmoke;
            _labelCounter.Location = new Point(14, 60);
            _labelCounter.Size = new Size(60, 20);

            // _buttonPrev
            _buttonPrev.FlatStyle = FlatStyle.Flat;
            _buttonPrev.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            _buttonPrev.ForeColor = Color.White;
            _buttonPrev.BackColor = Color.FromArgb(64, 64, 64);
            _buttonPrev.Location = new Point(940, 14);
            _buttonPrev.Size = new Size(60, 28);
            _buttonPrev.Text = "Prev";
            _buttonPrev.TabStop = false;

            // _buttonNext
            _buttonNext.FlatStyle = FlatStyle.Flat;
            _buttonNext.Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            _buttonNext.ForeColor = Color.White;
            _buttonNext.BackColor = Color.FromArgb(64, 64, 64);
            _buttonNext.Location = new Point(940, 44);
            _buttonNext.Size = new Size(60, 28);
            _buttonNext.Text = "Next";
            _buttonNext.TabStop = false;

            // _buttonDismiss
            _buttonDismiss.FlatStyle = FlatStyle.Flat;
            _buttonDismiss.Font = new Font("Segoe UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            _buttonDismiss.ForeColor = Color.White;
            _buttonDismiss.BackColor = Color.FromArgb(244, 67, 54);
            _buttonDismiss.Location = new Point(1008, 14);
            _buttonDismiss.Size = new Size(60, 56);
            _buttonDismiss.Text = "Dismiss";
            _buttonDismiss.TabStop = false;

            // TickerOverlayForm
            AutoScaleMode = AutoScaleMode.None;
            BackColor = Color.FromArgb(30, 30, 30);
            ClientSize = new Size(1085, 80);
            Controls.Add(_accentBar);
            Controls.Add(_labelMessage);
            Controls.Add(_labelCounter);
            Controls.Add(_buttonPrev);
            Controls.Add(_buttonNext);
            Controls.Add(_buttonDismiss);
            DoubleBuffered = true;
            FormBorderStyle = FormBorderStyle.None;
            Name = "TickerOverlayForm";
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            TopMost = true;
            ResumeLayout(false);
            PerformLayout();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }

            base.Dispose(disposing);
        }
    }
}
