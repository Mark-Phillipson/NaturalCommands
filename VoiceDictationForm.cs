using System;
using System.Drawing;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Text.Json;
using System.Collections.Generic;
using System.Linq;
using System.Media;
using WindowsInput;
using WindowsInput.Native;
using NaturalCommands.Helpers;
using NaturalCommands.Models;

namespace DictationBoxMSP
{
    public class VoiceDictationForm : Form
    {
        private TextBox txtInput = null!;
        private Button btnCancel = null!;
        // removed Restart button (use voice phrase "voice typing" to start)
        private Button btnSendCommand = null!;
        private Button btnCopyText = null!;
        private Button btnClearText = null!;
        private Button btnSearchWeb = null!;
        private Button btnOpenInVsc = null!;
        private Button btnSendToCopilot = null!; // new button for Copilot CLI
        private Button btnCallAssistant = null!; // new button for personal-assistant CLI
        private Button btnToggleTransparent = null!;
        private Panel bottomPanel = null!;
        private Label lblTransient = null!;
        private List<Label> marqueeLabels = new List<Label>();
        private System.Windows.Forms.Timer marqueeTimer = null!;
        private System.Windows.Forms.Timer autoSubmitTimer = null!;
        private System.Windows.Forms.Timer startDictationTimer = null!;
        private System.Windows.Forms.Timer dictationStopDebounceTimer = null!;
        private readonly DebounceGate dictationStopDebounce = new DebounceGate(850);
        // When true, detect dictation 'stopped' on short pauses. Default false to keep voice typing enabled until user stops it.
        private readonly bool autoDetectDictationStop = false;
        private readonly TemporaryMarqueeOverride temporaryMarqueeOverride = new TemporaryMarqueeOverride();
        private int timeoutMs = 0;
        private bool isBackgroundTransparent = false;
        // removed unused saved colors to avoid compiler warnings
        private Color savedTxtInputBackColor;
        private Color savedBottomPanelBackColor;
        private double savedOpacity = 1.0;

        // Command history for terminal-like up/down navigation
        private List<string> commandHistory = new List<string>();
        private int historyIndex = -1;
        private string currentInput = string.Empty; // Saves current input when navigating history

        public string ResultText => txtInput.Text ?? string.Empty;

        /// <summary>
        /// Event fired when user sends a command. The form stays open for additional commands.
        /// </summary>
        public event EventHandler<string>? CommandSent;

        public VoiceDictationForm(int timeoutMs = -1, bool autoStartDictation = true, bool autoDetectStop = false)
        {
            this.timeoutMs = timeoutMs;
            // Allow optional enabling of automatic 'dictation stopped' detection (defaults to false)
            try { autoDetectDictationStop = autoDetectStop; } catch { }
            InitializeComponents();
            ApplySharedStyles();

            // Apply transparent setting from AppSettings if enabled
            try
            {
                // Preserve values so Toggle can restore them later
                savedOpacity = this.Opacity;
                savedTxtInputBackColor = txtInput.BackColor;
                savedBottomPanelBackColor = bottomPanel.BackColor;

                // Check if user wants to start in transparent mode
                bool startTransparent = false;
                try { startTransparent = AppSettings.Instance.VoiceDictation.StartTransparent; } catch { }

                if (startTransparent)
                {
                    // Apply transparency immediately
                    this.Opacity = 0.65;
                    txtInput.BackColor = DisplayMessage.SharedBackColor;
                    bottomPanel.BackColor = DisplayMessage.SharedBackColor;
                    isBackgroundTransparent = true;
                    try { btnToggleTransparent.Text = "Disable Trans&parent"; } catch { }
                }
                else
                {
                    // Start opaque by default
                    isBackgroundTransparent = false;
                    try { btnToggleTransparent.Text = "Toggle Trans&parent"; } catch { }
                }
            }
            catch { }

            if (autoStartDictation && timeoutMs >= 0)
            {
                this.Shown += VoiceDictationForm_Shown;
            }
        }

        private void InitializeComponents()
        {
            this.txtInput = new TextBox() { Multiline = true, Dock = DockStyle.Fill, ScrollBars = ScrollBars.Vertical };
            // Use '&' to indicate keyboard accelerators (mnemonics). These are shown as underlined
            // when Alt is pressed and allow keyboard activation (Alt+Key).
            // Re-Start Dictation button removed (use voice phrase instead)
            this.btnCancel = new Button() { Text = "&Cancel", Height = 121 };
            // Use Alt+S for Send Command (replaces the removed Submit button)
            this.btnSendCommand = new Button() { Text = "&Send Command", Height = 121 };
            this.btnCopyText = new Button() { Text = "Copy &Text", Height = 121 };
            this.btnClearText = new Button() { Text = "Clear Text", Height = 121 };
            this.btnSearchWeb = new Button() { Text = "Search &Web", Height = 121 };
            this.btnOpenInVsc = new Button() { Text = "Open in &VS Code", Height = 121 };
            this.btnSendToCopilot = new Button() { Text = "Send to Copi&lot", Height = 121 }; // Copilot CLI
            this.btnCallAssistant = new Button() { Text = "Call &Assistant", Height = 121 }; // Personal Assistant CLI
            this.autoSubmitTimer = new System.Windows.Forms.Timer();
            this.startDictationTimer = new System.Windows.Forms.Timer();
            this.dictationStopDebounceTimer = new System.Windows.Forms.Timer();
            this.dictationStopDebounceTimer.Interval = 90;
            this.dictationStopDebounceTimer.Tick += DictationStopDebounceTimer_Tick;

            // Bottom area: use a table so marquee cannot overlap buttons.
            bottomPanel = new Panel() { Dock = DockStyle.Bottom, Height = 140 };
            bottomPanel.Padding = new Padding(8);
            bottomPanel.BackColor = DisplayMessage.SharedBackColor;

            // We'll use a 3-row TableLayout: marquee (36px), buttons (84px), transient (20px)
            var table = new TableLayoutPanel() { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
            table.RowStyles.Clear();
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 84));
            table.RowStyles.Add(new RowStyle(SizeType.Absolute, 20));

            // Marquee row (top of bottomPanel)
            var marqueePanel = new Panel() { Dock = DockStyle.Fill, Height = 28 };
            marqueePanel.Padding = new Padding(6, 6, 6, 6);
            marqueePanel.BackColor = DisplayMessage.SharedBackColor;

            // Create multiple marquee labels so we can show more than one message at once
            int marqueeCount = 2;
            for (int i = 0; i < marqueeCount; i++)
            {
                var lbl = new Label()
                {
                    AutoSize = true,
                    TextAlign = ContentAlignment.MiddleLeft,
                    ForeColor = DisplayMessage.SharedForeColor,
                    BackColor = Color.Transparent,
                    Font = new Font(this.Font.FontFamily, Math.Max(this.Font.Size - 6f, 10f), FontStyle.Regular)
                };
                marqueePanel.Controls.Add(lbl);
                marqueeLabels.Add(lbl);
            }

            // Buttons row (middle)
            var buttonsContainer = new Panel() { Dock = DockStyle.Fill };
            var flow = new FlowLayoutPanel() { Dock = DockStyle.Fill, FlowDirection = FlowDirection.LeftToRight, AutoSize = false };
            // allow buttons to wrap onto a second line when there isn't enough width
            flow.WrapContents = true;
            flow.Padding = new Padding(6);
            // shorten button heights so marquee and buttons are visible together
            var btnHeight = 64;
            this.btnCancel.Height = btnHeight;
            this.btnSendCommand.Height = btnHeight;
            this.btnCopyText.Height = btnHeight;
            this.btnSearchWeb.Height = btnHeight;
            this.btnOpenInVsc.Height = btnHeight;
            this.btnSendToCopilot.Height = btnHeight; // new button height
            this.btnCallAssistant.Height = btnHeight; // new button height
            this.btnClearText.Height = btnHeight;
            this.btnToggleTransparent = new Button() { Text = "Toggle Trans&parent", Height = btnHeight };

            flow.Controls.Add(btnCancel);
            flow.Controls.Add(btnSendCommand);
            flow.Controls.Add(btnCopyText);
            flow.Controls.Add(btnClearText);
            flow.Controls.Add(btnOpenInVsc);
            flow.Controls.Add(btnSendToCopilot); // add to flow layout
            flow.Controls.Add(btnCallAssistant); // add to flow layout
            flow.Controls.Add(btnToggleTransparent);
            flow.Controls.Add(btnSearchWeb);

            // Keep explicit widths for remaining buttons so text remains visible
            btnCancel.Width = 140; btnSendCommand.Width = 170; btnCopyText.Width = 160; btnClearText.Width = 140; btnSearchWeb.Width = 160; btnOpenInVsc.Width = 180; btnSendToCopilot.Width = 180; btnCallAssistant.Width = 160; btnToggleTransparent.Width = 180;

            // Make button fonts slightly larger so text is easier to read
                try
                {
                    var baseSize = Math.Max(this.Font.Size * 0.85f, 10f);
                    var btnFont = new Font(this.Font.FontFamily, baseSize + 1f, FontStyle.Regular);
                    btnCancel.Font = btnFont;
                    btnSendCommand.Font = btnFont;
                    btnCopyText.Font = btnFont;
                    btnClearText.Font = btnFont;
                    btnSearchWeb.Font = btnFont;
                    btnOpenInVsc.Font = btnFont;
                    btnSendToCopilot.Font = btnFont;
                    btnCallAssistant.Font = btnFont;
                    btnToggleTransparent.Font = btnFont;
                }
                catch { }

            buttonsContainer.Controls.Add(flow);

            // Fix tab order to match visual left-to-right order. With FlowDirection=RightToLeft the visual
            // left-to-right ordering of controls is the reverse of the addition order. Explicitly set TabIndex
            // so pressing Tab moves focus in the same order the buttons appear on screen.
            try
            {
                // Ensure text input receives focus first
                txtInput.TabIndex = 0;
                // Visual left-to-right order: Search Web, Toggle Transparent, Open in VS Code, Call Assistant, Copilot CLI, Copy Text, Clear Text, Send Command, Cancel
                btnSearchWeb.TabIndex = 1;
                btnToggleTransparent.TabIndex = 2;
                btnOpenInVsc.TabIndex = 3;
                btnCallAssistant.TabIndex = 4;
                btnSendToCopilot.TabIndex = 5;
                btnCopyText.TabIndex = 6;
                btnClearText.TabIndex = 7;
                btnSendCommand.TabIndex = 8;
                btnCancel.TabIndex = 9;
            }
            catch { }

            // Transient row (bottom of bottomPanel)
            lblTransient = new Label()
            {
                Dock = DockStyle.Fill,
                Height = 20,
                TextAlign = ContentAlignment.MiddleCenter,
                Visible = false,
                BackColor = Color.LimeGreen,
                ForeColor = Color.Black,
                Font = new Font(this.Font.FontFamily, Math.Max(this.Font.Size - 2f, 10f), FontStyle.Bold),
                Padding = new Padding(4)
            };

            table.Controls.Add(marqueePanel, 0, 0);
            table.Controls.Add(buttonsContainer, 0, 1);
            // lblTransient already created above and configured when building the table
            table.Controls.Add(lblTransient, 0, 2);

            bottomPanel.Controls.Add(table);

            // Start marquee behavior: scroll multiple labels across the panel
            try
            {
                var marqueeItems = LoadMarqueeItems();
                var rnd = new Random();
                marqueeTimer = new System.Windows.Forms.Timer();
                marqueeTimer.Interval = 28;
                marqueeTimer.Tick += (s, e) =>
                {
                    try
                    {
                        var nowMs = Environment.TickCount64;
                        foreach (var lbl in marqueeLabels.ToList())
                        {
                            if (lbl.Left + lbl.Width <= -10)
                            {
                                var text = temporaryMarqueeOverride.GetMessage(nowMs) ?? (marqueeItems.Count > 0 ? marqueeItems[rnd.Next(marqueeItems.Count)] : "Say 'voice typing' to begin");
                                lbl.Text = text;
                                // place this label to the right of the right-most label (or panel if none)
                                int rightmost = marqueePanel.Width;
                                foreach (var other in marqueeLabels)
                                {
                                    if (other == lbl) continue;
                                    rightmost = Math.Max(rightmost, other.Left + other.Width + 40);
                                }
                                lbl.Left = rightmost;
                                lbl.Top = (marqueePanel.Height - lbl.Height) / 2;
                            }
                            else
                            {
                                lbl.Left -= 2;
                            }
                        }
                    }
                    catch { }
                };

                this.Shown += (s, e) =>
                {
                    try
                    {
                        // Do not pre-set a long-lived temporary marquee here. Temporary marquee should only be shown
                        // when the user has focus on the Send Command button. Initialize marquee labels with normal items.
                        for (int i = 0; i < marqueeLabels.Count; i++)
                        {
                            var lbl = marqueeLabels[i];
                            if (string.IsNullOrEmpty(lbl.Text))
                            {
                                lbl.Text = marqueeItems.Count > 0 ? marqueeItems[rnd.Next(marqueeItems.Count)] : "Say 'voice typing' to begin";
                            }
                            // stagger starting positions so messages are spread out
                            lbl.Left = marqueePanel.Width + i * (marqueePanel.Width / 2);
                            lbl.Top = (marqueePanel.Height - lbl.Height) / 2;
                        }
                        marqueeTimer?.Start();
                    }
                    catch { }
                };
            }
            catch { }

            // Make button borders visible on all sides by using FlatStyle and a small border
            btnCancel.FlatStyle = FlatStyle.Flat; btnCancel.FlatAppearance.BorderSize = 1; btnCancel.FlatAppearance.BorderColor = SystemColors.ControlDark; btnCancel.Margin = new Padding(6);
            btnSendCommand.FlatStyle = FlatStyle.Flat; btnSendCommand.FlatAppearance.BorderSize = 1; btnSendCommand.FlatAppearance.BorderColor = SystemColors.ControlDark; btnSendCommand.Margin = new Padding(6);
            btnCopyText.FlatStyle = FlatStyle.Flat; btnCopyText.FlatAppearance.BorderSize = 1; btnCopyText.FlatAppearance.BorderColor = SystemColors.ControlDark; btnCopyText.Margin = new Padding(6);
            btnClearText.FlatStyle = FlatStyle.Flat; btnClearText.FlatAppearance.BorderSize = 1; btnClearText.FlatAppearance.BorderColor = SystemColors.ControlDark; btnClearText.Margin = new Padding(6);
            btnOpenInVsc.FlatStyle = FlatStyle.Flat; btnOpenInVsc.FlatAppearance.BorderSize = 1; btnOpenInVsc.FlatAppearance.BorderColor = SystemColors.ControlDark; btnOpenInVsc.Margin = new Padding(6);
            btnSendToCopilot.FlatStyle = FlatStyle.Flat; btnSendToCopilot.FlatAppearance.BorderSize = 1; btnSendToCopilot.FlatAppearance.BorderColor = SystemColors.ControlDark; btnSendToCopilot.Margin = new Padding(6);
            btnCallAssistant.FlatStyle = FlatStyle.Flat; btnCallAssistant.FlatAppearance.BorderSize = 1; btnCallAssistant.FlatAppearance.BorderColor = SystemColors.ControlDark; btnCallAssistant.Margin = new Padding(6);
            btnSearchWeb.FlatStyle = FlatStyle.Flat; btnSearchWeb.FlatAppearance.BorderSize = 1; btnSearchWeb.FlatAppearance.BorderColor = SystemColors.ControlDark; btnSearchWeb.Margin = new Padding(6);
            btnToggleTransparent.FlatStyle = FlatStyle.Flat; btnToggleTransparent.FlatAppearance.BorderSize = 1; btnToggleTransparent.FlatAppearance.BorderColor = SystemColors.ControlDark; btnToggleTransparent.Margin = new Padding(6);

            // Provide tooltips and ensure mnemonics are enabled so shortcuts are discoverable
            try
            {
                var tt = new ToolTip();
                tt.IsBalloon = false;
                tt.ShowAlways = true;
                tt.SetToolTip(btnToggleTransparent, "Toggle Transparent (Alt+P)");
                tt.SetToolTip(btnCopyText, "Copy text to clipboard (Alt+T)");
                tt.SetToolTip(btnClearText, "Clear text and focus the input");
                tt.SetToolTip(btnSendToCopilot, "Send text to Copilot CLI in a new terminal");
                tt.SetToolTip(btnCallAssistant, "Send text to Personal Assistant CLI in a new terminal (Alt+A)");
                // Ensure mnemonics are enabled explicitly
                btnToggleTransparent.UseMnemonic = true;
                btnCopyText.UseMnemonic = true;
                btnClearText.UseMnemonic = true;
                btnSendToCopilot.UseMnemonic = true;
                btnCallAssistant.UseMnemonic = true;
            }
            catch { }

            this.Controls.Add(txtInput);
            this.Controls.Add(bottomPanel);

            this.Text = "Voice Dictation";
            this.StartPosition = FormStartPosition.CenterScreen;
            // Make the form a little wider and twice as high so buttons don't get cut off
            // increasing width helps when many buttons are present
            this.Size = new Size(1400, 800);

            btnCancel.Click += BtnCancel_Click;
            // start button removed; dictation triggered via voice phrase
            btnSendCommand.Click += BtnSendCommand_Click;
            btnSendCommand.GotFocus += BtnSendCommand_GotFocus;
            btnSendCommand.LostFocus += BtnSendCommand_LostFocus;
            btnCopyText.Click += BtnCopyText_Click;
            btnClearText.Click += BtnClearText_Click;
            btnOpenInVsc.Click += BtnOpenInVsc_Click;
            btnSendToCopilot.Click += BtnSendToCopilot_Click; // copilot handler
            btnCallAssistant.Click += BtnCallAssistant_Click; // personal assistant handler
            btnToggleTransparent.Click += BtnToggleTransparent_Click;
            btnSearchWeb.Click += BtnSearchWeb_Click;
            this.FormClosing += VoiceDictationForm_FormClosing;
            this.KeyPreview = true;
            this.KeyDown += VoiceDictationForm_KeyDown;

            txtInput.TextChanged += TxtInput_TextChanged;
            txtInput.KeyDown += TxtInput_KeyDown;
        }

        private void ApplySharedStyles()
        {
            try
            {
                this.BackColor = DisplayMessage.SharedBackColor;
                this.ForeColor = DisplayMessage.SharedForeColor;
                var baseFont = DisplayMessage.SharedFont ?? SystemFonts.MessageBoxFont;
                var larger = new Font(baseFont!.FontFamily, Math.Max(baseFont!.Size * 1.8f, baseFont!.Size + 8f), baseFont!.Style);
                this.Font = larger;
                txtInput.Font = larger;
                txtInput.BackColor = DisplayMessage.SharedBackColor;
                txtInput.ForeColor = DisplayMessage.SharedForeColor;
                try { if (bottomPanel != null) bottomPanel.BackColor = DisplayMessage.SharedBackColor; } catch { }
            }
            catch { }
        }

        private void BtnToggleTransparent_Click(object? sender, EventArgs e)
        {
            try
            {
                if (!isBackgroundTransparent)
                {
                    // save current opacity and colors
                    savedOpacity = this.Opacity;
                    savedTxtInputBackColor = txtInput.BackColor;
                    savedBottomPanelBackColor = bottomPanel.BackColor;

                    // set a semi-transparent window (affects controls too)
                    this.Opacity = 0.65;

                    // keep control backgrounds readable
                    txtInput.BackColor = DisplayMessage.SharedBackColor;
                    bottomPanel.BackColor = DisplayMessage.SharedBackColor;

                    isBackgroundTransparent = true;
                    // Keep the same mnemonic (Trans&parent -> Alt+P) for the disabled state
                    btnToggleTransparent.Text = "Disable Trans&parent";
                }
                else
                {
                    // restore
                    this.Opacity = savedOpacity;
                    txtInput.BackColor = savedTxtInputBackColor;
                    bottomPanel.BackColor = savedBottomPanelBackColor;
                    isBackgroundTransparent = false;
                    btnToggleTransparent.Text = "Toggle Trans&parent";
                }
            }
            catch { }
        }

        private void VoiceDictationForm_Shown(object? sender, EventArgs e)
        {
            try
            {
                this.BringToFront();
                this.Activate();
                txtInput.Focus();
                txtInput.Select();
            }
            catch { }

            // Start dictation shortly after shown so focus is established
            startDictationTimer.Interval = 300;
            startDictationTimer.Tick += StartDictationTimer_Tick;
            startDictationTimer.Start();

            if (timeoutMs > 0)
            {
                autoSubmitTimer.Interval = timeoutMs;
                autoSubmitTimer.Tick += AutoSubmitTimer_Tick;
                autoSubmitTimer.Start();
            }
        }

        private void StartDictationTimer_Tick(object? sender, EventArgs e)
        {
            try { startDictationTimer.Stop(); } catch { }
            StartDictation();
        }

        // BtnStart removed — StartDictation can still be invoked by voice phrase or timers

        private void StartDictation()
        {
            try
            {
                txtInput.Focus();
                txtInput.Select();
                var sim = new InputSimulator();
                // Make sure modifier keys (Alt/Ctrl/Shift) are released before simulating Win+H.
                // When the user activates the button via an access key (Alt+Key) the Alt key
                // may still be logically down which can interfere with the Win+H keystroke.
                try
                {
                    sim.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.MENU);
                    sim.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.CONTROL);
                    sim.Keyboard.KeyUp(WindowsInput.Native.VirtualKeyCode.SHIFT);
                }
                catch { }

                // Give the input system a moment to settle after releasing modifiers
                System.Threading.Thread.Sleep(60);

                sim.Keyboard.ModifiedKeyStroke(VirtualKeyCode.LWIN, VirtualKeyCode.VK_H);
            }
            catch { }
        }

        private void AutoSubmitTimer_Tick(object? sender, EventArgs e)
        {
            autoSubmitTimer.Stop();
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void BtnCancel_Click(object? sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private async void BtnSendCommand_Click(object? sender, EventArgs e)
        {
            var commandText = txtInput.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrEmpty(commandText))
            {
                txtInput.Focus();
                return;
            }

            try
            {
                // Add to command history (avoid duplicating the last command)
                if (commandHistory.Count == 0 || commandHistory[commandHistory.Count - 1] != commandText)
                {
                    commandHistory.Add(commandText);
                }
                // Reset history navigation
                historyIndex = -1;
                currentInput = string.Empty;

                // Show visual feedback that command is being processed
                btnSendCommand.Enabled = false;
                var originalText = btnSendCommand.Text;
                btnSendCommand.Text = "Sending...";
                btnSendCommand.Refresh();

                // Show transient feedback
                if (lblTransient != null)
                {
                    lblTransient.Text = $"Sending: {(commandText.Length > 50 ? commandText.Substring(0, 50) + "..." : commandText)}";
                    lblTransient.BackColor = Color.DodgerBlue;
                    lblTransient.ForeColor = Color.White;
                    lblTransient.Visible = true;
                    lblTransient.BringToFront();
                }

                // Fire the event so the caller can process the command
                CommandSent?.Invoke(this, commandText);

                // Clear the text and refocus for the next command
                txtInput.Clear();
                txtInput.Focus();

                // Brief delay to show feedback
                await Task.Delay(300);

                // Update feedback to show command was sent
                if (lblTransient != null)
                {
                    lblTransient.Text = "Command sent! Ready for next command.";
                    lblTransient.BackColor = Color.LimeGreen;
                    lblTransient.ForeColor = Color.Black;
                }

                // Restore button
                btnSendCommand.Text = originalText;
                btnSendCommand.Enabled = true;

                // Hide feedback after a moment
                await Task.Delay(1500);
                if (lblTransient != null)
                    lblTransient.Visible = false;
            }
            catch
            {
                // On error, restore button state
                btnSendCommand.Text = "&Send Command";
                btnSendCommand.Enabled = true;
                txtInput.Focus();
            }
        }

        public void ShowTemporaryMarquee(string message, int durationMs)
        {
            try
            {
                temporaryMarqueeOverride.Set(message, durationMs, Environment.TickCount64);
            }
            catch { }

            try
            {
                foreach (var lbl in marqueeLabels)
                    lbl.Text = message;
            }
            catch { }
        }

        private void TxtInput_TextChanged(object? sender, EventArgs e)
        {
            try
            {
                if (lblTransient != null)
                    lblTransient.Visible = false;
            }
            catch { }

            try
            {
                if (autoDetectDictationStop)
                {
                    dictationStopDebounce.MarkChange(Environment.TickCount64);
                    if (!dictationStopDebounceTimer.Enabled)
                        dictationStopDebounceTimer.Start();
                }
            }
            catch { }
        }

        private void DictationStopDebounceTimer_Tick(object? sender, EventArgs e)
        {
            try
            {
                if (!dictationStopDebounce.TryConsume(Environment.TickCount64))
                    return;

                dictationStopDebounceTimer.Stop();
                OnDictationStopped();
            }
            catch { }
        }

        private async void OnDictationStopped()
        {
            try { if (!autoDetectDictationStop) return; } catch { }

            try { this.AcceptButton = btnSendCommand; } catch { }

            try
            {
                if (this.Visible && this.ContainsFocus)
                    btnSendCommand.Focus();
            }
            catch { }

            try
            {
                if (lblTransient != null)
                {
                    lblTransient.Text = "Listening finished — press Enter to send";
                    lblTransient.BackColor = Color.LimeGreen;
                    lblTransient.ForeColor = Color.Black;
                    lblTransient.Visible = true;
                    lblTransient.BringToFront();
                }
            }
            catch { }

            try { SystemSounds.Asterisk.Play(); } catch { }

            try
            {
                await Task.Delay(2000);
                if (lblTransient != null) lblTransient.Visible = false;
            }
            catch { }
        }

        private async void BtnCopyText_Click(object? sender, EventArgs e)
        {
            try
            {
                var text = txtInput.Text ?? string.Empty;
                if (string.IsNullOrEmpty(text)) return;

                Clipboard.SetText(text);

                try
                {
                    // Stronger visual feedback: change text, colors and force a refresh so it's obvious
                    var originalText = btnCopyText.Text;
                    var originalBack = btnCopyText.BackColor;
                    var originalFore = btnCopyText.ForeColor;
                    var originalFont = btnCopyText.Font;

                    btnCopyText.Enabled = false;
                    btnCopyText.Text = "Copied";
                    // Ensure BackColor will be applied even when visual styles are enabled
                    var originalUseVisual = btnCopyText.UseVisualStyleBackColor;
                    try { btnCopyText.UseVisualStyleBackColor = false; } catch { }
                    btnCopyText.BackColor = Color.LimeGreen;
                    btnCopyText.ForeColor = Color.Black;
                    btnCopyText.Font = new Font(originalFont.FontFamily, originalFont.Size, FontStyle.Bold);
                    btnCopyText.Refresh();

                    // Show the transient label as a very visible confirmation below the buttons
                    try
                    {
                        if (lblTransient != null)
                        {
                            lblTransient.Text = "Copied";
                            lblTransient.BackColor = Color.LimeGreen;
                            lblTransient.ForeColor = Color.Black;
                            lblTransient.Visible = true;
                            lblTransient.BringToFront();
                        }
                    }
                    catch { }

                    await Task.Delay(1250);

                    btnCopyText.Text = originalText;
                    btnCopyText.BackColor = originalBack;
                    btnCopyText.ForeColor = originalFore;
                    btnCopyText.Font = originalFont;
                    try { btnCopyText.UseVisualStyleBackColor = originalUseVisual; } catch { }
                    btnCopyText.Enabled = true;
                    btnCopyText.Refresh();

                    // Hide the transient label after restoring button state
                    try
                    {
                        if (lblTransient != null)
                        {
                            lblTransient.Visible = false;
                        }
                    }
                    catch { }
                }
                catch { }

                try { NaturalCommands.TrayNotificationHelper.ShowNotification("Copied", "Text copied to clipboard", 1200); } catch { }
            }
            catch { }
        }

        private void BtnClearText_Click(object? sender, EventArgs e)
        {
            try
            {
                txtInput.Clear();
                txtInput.Focus();
            }
            catch { }
        }

        private void BtnSearchWeb_Click(object? sender, EventArgs e)
        {
            try
            {
                var text = txtInput.Text ?? string.Empty;
                if (string.IsNullOrWhiteSpace(text)) return;
                var query = Uri.EscapeDataString(text);
                var url = $"https://www.bing.com/search?q={query}";

                var psi = new ProcessStartInfo
                {
                    FileName = "msedge",
                    Arguments = url,
                    UseShellExecute = true
                };

                try
                {
                    Process.Start(psi);
                }
                catch
                {
                    // Fallback: open with default browser
                    var psi2 = new ProcessStartInfo
                    {
                        FileName = url,
                        UseShellExecute = true
                    };
                    Process.Start(psi2);
                }
            }
            catch { }
        }

        private void BtnSendToCopilot_Click(object? sender, EventArgs e)
        {
            try
            {
                var text = txtInput.Text ?? string.Empty;
                if (string.IsNullOrWhiteSpace(text)) return;

                // copy text to clipboard so user can paste elsewhere if desired
                try { Clipboard.SetText(text); } catch { }

                bool hasCopilot = IsCommandAvailable("copilot");
                bool hasGh = IsCommandAvailable("gh");
                if (!hasCopilot && !hasGh)
                {
                    MessageBox.Show("GitHub Copilot CLI not found on PATH.\nPlease install 'copilot' or the GitHub CLI with the Copilot extension.",
                        "Copilot CLI Missing", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Determine the command string, omitting text when multiline since CLI handles it poorly
                var cliCmd = GetCopilotCommand(text, hasCopilot, hasGh);
                LaunchTerminalWithCommand(cliCmd);

                // update transient label for user feedback
                if (lblTransient != null)
                {
                    lblTransient.Text = "Copied to clipboard and opened Copilot CLI";
                    lblTransient.BackColor = Color.LimeGreen;
                    lblTransient.ForeColor = Color.Black;
                    lblTransient.Visible = true;
                    lblTransient.BringToFront();
                    Task.Run(async () =>
                    {
                        await Task.Delay(1500);
                        try { if (lblTransient != null) lblTransient.Visible = false; } catch { }
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to launch Copilot CLI: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnCallAssistant_Click(object? sender, EventArgs e)
        {
            try
            {
                var text = txtInput.Text ?? string.Empty;
                if (string.IsNullOrWhiteSpace(text)) return;

                // copy text to clipboard so user can paste elsewhere if desired
                try { Clipboard.SetText(text); } catch { }

                // Locate personal-assistant.exe
                var assistantPath = FindPersonalAssistantExecutable();
                if (string.IsNullOrWhiteSpace(assistantPath))
                {
                    MessageBox.Show("Personal Assistant executable not found.\nPlease ensure personal-assistant.exe is available on PATH or in the application directory.",
                        "Personal Assistant Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Build the command line
                var command = GetAssistantCommand(text, assistantPath);
                LaunchTerminalWithCommand(command);

                // update transient label for user feedback
                if (lblTransient != null)
                {
                    lblTransient.Text = "Copied to clipboard and opened Personal Assistant";
                    lblTransient.BackColor = Color.LimeGreen;
                    lblTransient.ForeColor = Color.Black;
                    lblTransient.Visible = true;
                    lblTransient.BringToFront();
                    Task.Run(async () =>
                    {
                        await Task.Delay(1500);
                        try { if (lblTransient != null) lblTransient.Visible = false; } catch { }
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to launch Personal Assistant: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string? FindPersonalAssistantExecutable()
        {
            try
            {
                // Try to find in PATH
                var psi = new ProcessStartInfo
                {
                    FileName = "where",
                    Arguments = "personal-assistant",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using var proc = Process.Start(psi);
                if (proc != null)
                {
                    proc.WaitForExit(1000);
                    if (proc.ExitCode == 0)
                    {
                        var output = proc.StandardOutput.ReadToEnd().Trim();
                        var lines = output.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                        if (lines.Length > 0 && File.Exists(lines[0]))
                            return lines[0];
                    }
                }
            }
            catch { }

            // Try common locations
            var possibleLocationPatterns = new[]
            {
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "personal-assistant.exe"),
                // Try sibling project paths (relative from NaturalCommands output directory)
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "personal-assistant", "bin", "Debug", "net10.0", "personal-assistant.exe"),
                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "personal-assistant", "bin", "Release", "net10.0", "publish", "personal-assistant.exe"),
                // Try absolute path
                @"C:\Users\MPhil\source\repos\personal-assistant\bin\Debug\net10.0\personal-assistant.exe",
                // Try nget installation
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".nget", "bin", "personal-assistant.exe"),
            };

            foreach (var path in possibleLocationPatterns)
            {
                try
                {
                    var expandedPath = Path.GetFullPath(path);
                    if (File.Exists(expandedPath))
                        return expandedPath;
                }
                catch { }
            }

            return null;
        }

        private string GetAssistantCommand(string text, string assistantPath)
        {
            // Build command: personal-assistant.exe --cli "prompt"
            // Escape quotes in the text
            var escapedText = text.Replace("\"", "\\\"");
            return $"\"{assistantPath}\" --cli \"{escapedText}\"";
        }

        private bool IsCommandAvailable(string command)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = "where",
                    Arguments = command,
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using var proc = Process.Start(psi);
                if (proc == null) return false;
                proc.WaitForExit(1000);
                return proc.ExitCode == 0;
            }
            catch { return false; }
        }

        // Build the command line that will be passed to the terminal.
        // If the supplied text contains newlines we avoid embedding it as an argument
        // because Copilot CLI can't handle multiline prompts well. In that case we
        // still copy the text to the clipboard so the user may paste it manually.
        // Always launch the CLI without passing the text as an argument.  We have already
        // copied the text to the clipboard prior to calling this method, so the user may
        // manually paste it into the Copilot session.
        private string GetCopilotCommand(string text, bool hasCopilot, bool hasGh)
        {
            if (hasCopilot) return "copilot";
            else return "gh copilot chat";
        }

        private void LaunchTerminalWithCommand(string command)
        {
            try
            {
                // Try Windows Terminal (wt.exe) first
                var wtPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "WindowsApps", "wt.exe");
                if (File.Exists(wtPath))
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = wtPath,
                        Arguments = $"cmd /k {command}",
                        UseShellExecute = true
                    };
                    Process.Start(psi);
                }
                else
                {
                    // fallback to classic cmd
                    var psi = new ProcessStartInfo
                    {
                        FileName = "cmd.exe",
                        Arguments = $"/c start cmd /k {command}",
                        UseShellExecute = true
                    };
                    Process.Start(psi);
                }
            }
            catch { }
        }

        private void BtnSendCommand_GotFocus(object? sender, EventArgs e)
        {
            try
            {
                // Only show the temporary marquee while the Send Command button has focus
                ShowTemporaryMarquee("Listening finished — press Enter to send", 30 * 1000);
            }
            catch { }
        }

        private void BtnSendCommand_LostFocus(object? sender, EventArgs e)
        {
            try
            {
                // Clear any temporary marquee so normal marquee items resume
                temporaryMarqueeOverride.Clear();
            }
            catch { }
        }

        private void VoiceDictationForm_FormClosing(object? sender, FormClosingEventArgs e)
        {
            try { autoSubmitTimer.Stop(); } catch { }
            try { dictationStopDebounceTimer.Stop(); } catch { }
        }

        private void VoiceDictationForm_KeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
            }
        }

        private void TxtInput_KeyDown(object? sender, KeyEventArgs e)
        {
            try
            {
                // Handle Up arrow - go back in history
                if (e.KeyCode == Keys.Up && !e.Alt && !e.Control)
                {
                    if (commandHistory.Count == 0) return;

                    // If we're at the end (not browsing history yet), save current input
                    if (historyIndex == -1)
                    {
                        currentInput = txtInput.Text ?? string.Empty;
                        historyIndex = commandHistory.Count - 1;
                    }
                    else if (historyIndex > 0)
                    {
                        historyIndex--;
                    }

                    if (historyIndex >= 0 && historyIndex < commandHistory.Count)
                    {
                        txtInput.Text = commandHistory[historyIndex];
                        txtInput.SelectionStart = txtInput.Text.Length;
                    }
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
                // Handle Down arrow - go forward in history
                else if (e.KeyCode == Keys.Down && !e.Alt && !e.Control)
                {
                    if (commandHistory.Count == 0 || historyIndex == -1) return;

                    if (historyIndex < commandHistory.Count - 1)
                    {
                        historyIndex++;
                        txtInput.Text = commandHistory[historyIndex];
                        txtInput.SelectionStart = txtInput.Text.Length;
                    }
                    else
                    {
                        // We've gone past the end, restore the current input
                        historyIndex = -1;
                        txtInput.Text = currentInput;
                        txtInput.SelectionStart = txtInput.Text.Length;
                    }
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
            }
            catch { }
        }

        private List<string> LoadMarqueeItems()
        {
            // Curated United Kingdom English voice-typing commands (UK variants only)
            return new List<string>
            {
                "Punctuation: say 'full stop' Note this will also stop Voice Typing and go back to Talon Voice",
                "New line: say 'new line'",
                "New paragraph: say 'new paragraph'",
                "Open quote / Close quote: say 'open quote' / 'close quote'",
                "Colon / Semicolon: say 'colon' / 'semicolon'",
                "Ellipsis: say 'ellipsis'",
                "Open parenthesis / Close parenthesis: say 'open parenthesis' / 'close parenthesis'",
                "Delete last spoken word/phrase: say 'delete that' to remove last phrase",
                "Select last spoken word or phrase: say 'select that'",
                "Press enter: say 'press enter'",
                "Press Backspace: say 'press backspace'",
                "Press Tab: say 'press tab'",
                "Press Space: say 'press space'",
                "Say 'voice typing' to open Windows Voice Typing and set Talon Voice to sleep",
                "Click 'Send to Copi\u00A0lot' or press Alt+L to open a terminal and run Copilot CLI with the text",
                "Click 'Call &Assistant' or press Alt+A to open a terminal and run Personal Assistant CLI with the text"
            };
        }

        private void BtnOpenInVsc_Click(object? sender, EventArgs e)
        {
            try
            {
                var text = txtInput.Text ?? string.Empty;
                if (string.IsNullOrWhiteSpace(text)) return;

                var tempDir = Path.GetTempPath();
                var fileName = $"dictation-{DateTime.Now:yyyyMMdd-HHmmss}.txt";
                var filePath = Path.Combine(tempDir, fileName);
                File.WriteAllText(filePath, text);

                var args = $"--new-window \"{filePath}\"";

                // Try the 'code' CLI first
                try
                {
                    var psi = new ProcessStartInfo
                    {
                        FileName = "code",
                        Arguments = args,
                        UseShellExecute = true
                    };
                    Process.Start(psi);
                    return;
                }
                catch { }

                // Try common install locations
                var possible = new[]
                {
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft VS Code", "Code.exe"),
                    Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft VS Code", "Code.exe")
                };

                foreach (var p in possible)
                {
                    if (File.Exists(p))
                    {
                        var psi2 = new ProcessStartInfo
                        {
                            FileName = p,
                            Arguments = args,
                            UseShellExecute = true
                        };
                        Process.Start(psi2);
                        return;
                    }
                }

                // Fallback: open with default program
                Process.Start(new ProcessStartInfo { FileName = filePath, UseShellExecute = true });
            }
            catch { }
        }
    }
}
