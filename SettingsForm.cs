using System;
using System.Drawing;
using System.Windows.Forms;
using NaturalCommands.Models;
using NaturalCommands.Helpers;
using DictationBoxMSP;

namespace NaturalCommands
{
    /// <summary>
    /// Comprehensive settings form for configuring all aspects of NaturalCommands.
    /// Accessible from the system tray icon.
    /// </summary>
    public class SettingsForm : Form
    {
        private TabControl tabControl = null!;
        private Button btnSave = null!;
        private Button btnCancel = null!;
        private Button btnApply = null!;
        
        // General Tab
        private CheckBox chkEnableNotifications = null!;
        private NumericUpDown numNotificationDuration = null!;
        private CheckBox chkDebugMode = null!;
        private CheckBox chkShowDebugDialog = null!;
        
        // Auto-Click Tab
        private NumericUpDown numAutoClickDelay = null!;
        private CheckBox chkAutoClickShowOverlay = null!;
        
        // Mouse Movement Tab
        private NumericUpDown numMouseDefaultSpeed = null!;
        private NumericUpDown numMouseMinSpeed = null!;
        private NumericUpDown numMouseMaxSpeed = null!;
        private NumericUpDown numMouseSpeedStep = null!;
        private NumericUpDown numMouseUpdateInterval = null!;
        
        // Voice Dictation Tab
        private NumericUpDown numVoiceAutoSubmitTimeout = null!;
        private CheckBox chkVoiceAutoDetectStop = null!;
        private NumericUpDown numVoiceStopDebounce = null!;
        private NumericUpDown numVoiceFormWidth = null!;
        private NumericUpDown numVoiceFormHeight = null!;
        private CheckBox chkVoiceStartTransparent = null!;
        private NumericUpDown numVoiceMarqueeInterval = null!;
        
        // Appearance Tab
        private Button btnBackgroundColor = null!;
        private Button btnForegroundColor = null!;
        private ComboBox cmbFontFamily = null!;
        private NumericUpDown numFontSize = null!;
        private ComboBox cmbRichTextFont = null!;
        private NumericUpDown numRichTextFontSize = null!;
        private NumericUpDown numAccessibilityMultiplier = null!;
        
        // AI Integration Tab
        private Label lblApiKeyStatus = null!;
        private Button btnConfigureApiKey = null!;
        private TextBox txtModelName = null!;
        private CheckBox chkEnableAIFallback = null!;
        private TextBox txtPromptFile = null!;
        private Button btnBrowsePromptFile = null!;
        
        // Advanced Tab
        private CheckBox chkLoggingEnabled = null!;
        private ComboBox cmbLogLevel = null!;
        private CheckBox chkClearLogOnStartup = null!;
        private NumericUpDown numMaxLogSizeMB = null!;
        private NumericUpDown numMultiActionDelay = null!;
        private CheckBox chkContinueOnError = null!;
        private NumericUpDown numMessageBoxTimeout = null!;
        private NumericUpDown numVSCommandChordDelay = null!;
        
        // File Paths Tab
        private Button btnOpenWordReplacements = null!;
        private Button btnOpenMultiActions = null!;
        private Button btnOpenEmojiMappings = null!;
        private Button btnOpenVSCommands = null!;
        private Button btnReloadConfigs = null!;
        private Button btnOpenSettingsFolder = null!;

        public SettingsForm()
        {
            InitializeForm();
            CreateTabs();
            LoadSettings();
        }

        private void InitializeForm()
        {
            Text = "NaturalCommands Settings";
            Size = new Size(700, 600);
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            
            // Apply shared styling
            Font = DisplayMessage.SharedFont;
            BackColor = DisplayMessage.SharedBackColor;
            ForeColor = DisplayMessage.SharedForeColor;
            
            // Create main layout
            var mainPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                Padding = new Padding(10)
            };
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            mainPanel.RowStyles.Add(new RowStyle(SizeType.Absolute, 50F));
            
            // Create tab control
            tabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = DisplayMessage.SharedFont,
                BackColor = DisplayMessage.SharedBackColor,
                ForeColor = DisplayMessage.SharedForeColor
            };
            
            // Create button panel
            var buttonPanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(0, 10, 0, 0),
                BackColor = DisplayMessage.SharedBackColor
            };
            
            btnCancel = CreateButton("Cancel", (s, e) => { DialogResult = DialogResult.Cancel; Close(); });
            btnApply = CreateButton("Apply", BtnApply_Click);
            btnSave = CreateButton("Save", BtnSave_Click);
            
            buttonPanel.Controls.Add(btnCancel);
            buttonPanel.Controls.Add(btnApply);
            buttonPanel.Controls.Add(btnSave);
            
            mainPanel.Controls.Add(tabControl, 0, 0);
            mainPanel.Controls.Add(buttonPanel, 0, 1);
            
            Controls.Add(mainPanel);
        }

        private Button CreateButton(string text, EventHandler clickHandler)
        {
            var btn = new Button
            {
                Text = text,
                Size = new Size(80, 30),
                FlatStyle = FlatStyle.Flat,
                Font = DisplayMessage.SharedFont
            };
            btn.FlatAppearance.BorderSize = 1;
            btn.Click += clickHandler;
            return btn;
        }

        private void CreateTabs()
        {
            tabControl.TabPages.Add(CreateGeneralTab());
            tabControl.TabPages.Add(CreateAutoClickTab());
            tabControl.TabPages.Add(CreateMouseMovementTab());
            tabControl.TabPages.Add(CreateVoiceDictationTab());
            tabControl.TabPages.Add(CreateAppearanceTab());
            tabControl.TabPages.Add(CreateAITab());
            tabControl.TabPages.Add(CreateAdvancedTab());
            tabControl.TabPages.Add(CreateFilePathsTab());
        }

        private TabPage CreateGeneralTab()
        {
            var tab = new TabPage("General");
            tab.BackColor = DisplayMessage.SharedBackColor;
            tab.ForeColor = DisplayMessage.SharedForeColor;
            var panel = CreateScrollablePanel();
            int yPos = 10;
            
            // Notifications
            panel.Controls.Add(CreateLabel("Notifications", ref yPos, true));
            chkEnableNotifications = CreateCheckBox("Enable notifications", ref yPos);
            panel.Controls.Add(chkEnableNotifications);
            panel.Controls.Add(CreateLabel("Notification duration (ms):", ref yPos));
            numNotificationDuration = CreateNumericUpDown(1000, 30000, ref yPos);
            panel.Controls.Add(numNotificationDuration);
            
            yPos += 20;
            
            // Debug
            panel.Controls.Add(CreateLabel("Debug Options", ref yPos, true));
            chkDebugMode = CreateCheckBox("Enable debug mode", ref yPos);
            panel.Controls.Add(chkDebugMode);
            chkShowDebugDialog = CreateCheckBox("Show debug dialog on startup", ref yPos);
            panel.Controls.Add(chkShowDebugDialog);
            
            tab.Controls.Add(panel);
            return tab;
        }

        private TabPage CreateAutoClickTab()
        {
            var tab = new TabPage("Auto-Click");
            tab.BackColor = DisplayMessage.SharedBackColor;
            tab.ForeColor = DisplayMessage.SharedForeColor;
            var panel = CreateScrollablePanel();
            int yPos = 10;
            
            panel.Controls.Add(CreateLabel("Auto-Click Settings", ref yPos, true));
            panel.Controls.Add(CreateLabel("Idle delay before click (ms):", ref yPos));
            numAutoClickDelay = CreateNumericUpDown(100, 2000, ref yPos);
            panel.Controls.Add(numAutoClickDelay);
            panel.Controls.Add(CreateLabel("Range: 100-2000ms. Default: 2000ms", ref yPos, false, 8));
            
            yPos += 10;
            chkAutoClickShowOverlay = CreateCheckBox("Show countdown overlay", ref yPos);
            panel.Controls.Add(chkAutoClickShowOverlay);
            
            yPos += 20;
            panel.Controls.Add(CreateLabel("Voice Commands:", ref yPos, true));
            panel.Controls.Add(CreateLabel("• 'auto click' - Enable auto-click mode", ref yPos, false, 9));
            panel.Controls.Add(CreateLabel("• 'stop auto click' - Disable auto-click mode", ref yPos, false, 9));
            
            tab.Controls.Add(panel);
            return tab;
        }

        private TabPage CreateMouseMovementTab()
        {
            var tab = new TabPage("Mouse Control");
            tab.BackColor = DisplayMessage.SharedBackColor;
            tab.ForeColor = DisplayMessage.SharedForeColor;
            var panel = CreateScrollablePanel();
            int yPos = 10;
            
            panel.Controls.Add(CreateLabel("Mouse Movement Speed", ref yPos, true));
            panel.Controls.Add(CreateLabel("Default speed (pixels/tick):", ref yPos));
            numMouseDefaultSpeed = CreateNumericUpDown(2, 50, ref yPos);
            panel.Controls.Add(numMouseDefaultSpeed);
            
            panel.Controls.Add(CreateLabel("Minimum speed:", ref yPos));
            numMouseMinSpeed = CreateNumericUpDown(1, 50, ref yPos);
            panel.Controls.Add(numMouseMinSpeed);
            
            panel.Controls.Add(CreateLabel("Maximum speed:", ref yPos));
            numMouseMaxSpeed = CreateNumericUpDown(2, 100, ref yPos);
            panel.Controls.Add(numMouseMaxSpeed);
            
            panel.Controls.Add(CreateLabel("Speed adjustment step:", ref yPos));
            numMouseSpeedStep = CreateNumericUpDown(1, 20, ref yPos);
            panel.Controls.Add(numMouseSpeedStep);
            
            panel.Controls.Add(CreateLabel("Update interval (ms):", ref yPos));
            numMouseUpdateInterval = CreateNumericUpDown(10, 100, ref yPos);
            panel.Controls.Add(numMouseUpdateInterval);
            panel.Controls.Add(CreateLabel("Lower = smoother (default: 16ms = ~60 FPS)", ref yPos, false, 8));
            
            tab.Controls.Add(panel);
            return tab;
        }

        private TabPage CreateVoiceDictationTab()
        {
            var tab = new TabPage("Voice Dictation");
            tab.BackColor = DisplayMessage.SharedBackColor;
            tab.ForeColor = DisplayMessage.SharedForeColor;
            var panel = CreateScrollablePanel();
            int yPos = 10;
            
            panel.Controls.Add(CreateLabel("Dictation Behavior", ref yPos, true));
            panel.Controls.Add(CreateLabel("Auto-submit timeout (ms, 0=disabled):", ref yPos));
            numVoiceAutoSubmitTimeout = CreateNumericUpDown(0, 60000, ref yPos);
            panel.Controls.Add(numVoiceAutoSubmitTimeout);
            
            chkVoiceAutoDetectStop = CreateCheckBox("Auto-detect dictation stop", ref yPos);
            panel.Controls.Add(chkVoiceAutoDetectStop);
            
            panel.Controls.Add(CreateLabel("Stop debounce delay (ms):", ref yPos));
            numVoiceStopDebounce = CreateNumericUpDown(100, 5000, ref yPos);
            panel.Controls.Add(numVoiceStopDebounce);
            
            panel.Controls.Add(CreateLabel("Marquee scroll interval (ms):", ref yPos));
            numVoiceMarqueeInterval = CreateNumericUpDown(10, 100, ref yPos);
            panel.Controls.Add(numVoiceMarqueeInterval);
            
            yPos += 20;
            panel.Controls.Add(CreateLabel("Form Appearance", ref yPos, true));
            panel.Controls.Add(CreateLabel("Default width:", ref yPos));
            numVoiceFormWidth = CreateNumericUpDown(400, 2000, ref yPos);
            panel.Controls.Add(numVoiceFormWidth);
            
            panel.Controls.Add(CreateLabel("Default height:", ref yPos));
            numVoiceFormHeight = CreateNumericUpDown(300, 1500, ref yPos);
            panel.Controls.Add(numVoiceFormHeight);
            
            chkVoiceStartTransparent = CreateCheckBox("Start in transparent mode", ref yPos);
            panel.Controls.Add(chkVoiceStartTransparent);
            
            tab.Controls.Add(panel);
            return tab;
        }

        private TabPage CreateAppearanceTab()
        {
            var tab = new TabPage("Appearance");
            tab.BackColor = DisplayMessage.SharedBackColor;
            tab.ForeColor = DisplayMessage.SharedForeColor;
            var panel = CreateScrollablePanel();
            int yPos = 10;
            
            panel.Controls.Add(CreateLabel("Colors", ref yPos, true));
            
            btnBackgroundColor = CreateButton("Select Background Color", (s, e) => PickColor(ref btnBackgroundColor));
            btnBackgroundColor.Location = new Point(20, yPos);
            btnBackgroundColor.Width = 200;
            panel.Controls.Add(btnBackgroundColor);
            yPos += 35;
            
            btnForegroundColor = CreateButton("Select Text Color", (s, e) => PickColor(ref btnForegroundColor));
            btnForegroundColor.Location = new Point(20, yPos);
            btnForegroundColor.Width = 200;
            panel.Controls.Add(btnForegroundColor);
            yPos += 45;
            
            panel.Controls.Add(CreateLabel("Fonts", ref yPos, true));
            panel.Controls.Add(CreateLabel("Font family:", ref yPos));
            cmbFontFamily = CreateFontComboBox(ref yPos);
            panel.Controls.Add(cmbFontFamily);
            
            panel.Controls.Add(CreateLabel("Font size (pt):", ref yPos));
            numFontSize = CreateNumericUpDown(6, 24, ref yPos, true);
            panel.Controls.Add(numFontSize);
            
            panel.Controls.Add(CreateLabel("Rich text font:", ref yPos));
            cmbRichTextFont = CreateFontComboBox(ref yPos);
            panel.Controls.Add(cmbRichTextFont);
            
            panel.Controls.Add(CreateLabel("Rich text font size (pt):", ref yPos));
            numRichTextFontSize = CreateNumericUpDown(8, 32, ref yPos, true);
            panel.Controls.Add(numRichTextFontSize);
            
            panel.Controls.Add(CreateLabel("Accessibility font multiplier:", ref yPos));
            numAccessibilityMultiplier = CreateNumericUpDown(1.0m, 3.0m, ref yPos, true, 0.1m);
            panel.Controls.Add(numAccessibilityMultiplier);
            
            tab.Controls.Add(panel);
            return tab;
        }

        private TabPage CreateAITab()
        {
            var tab = new TabPage("AI Integration");
            tab.BackColor = DisplayMessage.SharedBackColor;
            tab.ForeColor = DisplayMessage.SharedForeColor;
            var panel = CreateScrollablePanel();
            int yPos = 10;
            
            panel.Controls.Add(CreateLabel("OpenAI Configuration", ref yPos, true));
            
            lblApiKeyStatus = CreateLabel("API Key: Not configured", ref yPos);
            lblApiKeyStatus.ForeColor = Color.Orange;
            
            btnConfigureApiKey = CreateButton("Configure API Key", (s, e) => ConfigureApiKey());
            btnConfigureApiKey.Location = new Point(20, yPos);
            btnConfigureApiKey.Width = 180;
            panel.Controls.Add(btnConfigureApiKey);
            yPos += 45;
            
            panel.Controls.Add(CreateLabel("Model name:", ref yPos));
            txtModelName = CreateTextBox(ref yPos);
            panel.Controls.Add(txtModelName);
            
            chkEnableAIFallback = CreateCheckBox("Enable AI fallback for command interpretation", ref yPos);
            panel.Controls.Add(chkEnableAIFallback);
            
            yPos += 20;
            panel.Controls.Add(CreateLabel("Prompt file (relative to app directory):", ref yPos));
            txtPromptFile = CreateTextBox(ref yPos);
            panel.Controls.Add(txtPromptFile);
            
            btnBrowsePromptFile = CreateButton("Browse...", (s, e) => BrowsePromptFile());
            btnBrowsePromptFile.Location = new Point(420, yPos - 30);
            btnBrowsePromptFile.Width = 80;
            panel.Controls.Add(btnBrowsePromptFile);
            
            tab.Controls.Add(panel);
            return tab;
        }

        private TabPage CreateAdvancedTab()
        {
            var tab = new TabPage("Advanced");
            tab.BackColor = DisplayMessage.SharedBackColor;
            tab.ForeColor = DisplayMessage.SharedForeColor;
            var panel = CreateScrollablePanel();
            int yPos = 10;
            
            panel.Controls.Add(CreateLabel("Logging", ref yPos, true));
            chkLoggingEnabled = CreateCheckBox("Enable logging", ref yPos);
            panel.Controls.Add(chkLoggingEnabled);
            
            panel.Controls.Add(CreateLabel("Log level:", ref yPos));
            cmbLogLevel = CreateComboBox(new[] { "Debug", "Info", "Warning", "Error" }, ref yPos);
            panel.Controls.Add(cmbLogLevel);
            
            chkClearLogOnStartup = CreateCheckBox("Clear log on startup", ref yPos);
            panel.Controls.Add(chkClearLogOnStartup);
            
            panel.Controls.Add(CreateLabel("Max log file size (MB, 0=unlimited):", ref yPos));
            numMaxLogSizeMB = CreateNumericUpDown(0, 100, ref yPos);
            panel.Controls.Add(numMaxLogSizeMB);
            
            yPos += 20;
            panel.Controls.Add(CreateLabel("Multi-Action Settings", ref yPos, true));
            panel.Controls.Add(CreateLabel("Default delay between actions (ms):", ref yPos));
            numMultiActionDelay = CreateNumericUpDown(0, 5000, ref yPos);
            panel.Controls.Add(numMultiActionDelay);
            
            chkContinueOnError = CreateCheckBox("Continue on error", ref yPos);
            panel.Controls.Add(chkContinueOnError);
            
            yPos += 20;
            panel.Controls.Add(CreateLabel("Timing", ref yPos, true));
            panel.Controls.Add(CreateLabel("Message box auto-close timeout (ms):", ref yPos));
            numMessageBoxTimeout = CreateNumericUpDown(1000, 30000, ref yPos);
            panel.Controls.Add(numMessageBoxTimeout);
            
            panel.Controls.Add(CreateLabel("VS command chord delay (ms):", ref yPos));
            numVSCommandChordDelay = CreateNumericUpDown(50, 500, ref yPos);
            panel.Controls.Add(numVSCommandChordDelay);
            
            tab.Controls.Add(panel);
            return tab;
        }

        private TabPage CreateFilePathsTab()
        {
            var tab = new TabPage("File Paths");
            tab.BackColor = DisplayMessage.SharedBackColor;
            tab.ForeColor = DisplayMessage.SharedForeColor;
            var panel = CreateScrollablePanel();
            int yPos = 10;
            
            panel.Controls.Add(CreateLabel("Configuration Files", ref yPos, true));
            panel.Controls.Add(CreateLabel("Quick access to edit configuration files:", ref yPos));
            
            yPos += 10;
            
            btnOpenWordReplacements = CreateButton("Edit Word Replacements", (s, e) => OpenConfigFile("word_replacements.json"));
            btnOpenWordReplacements.Location = new Point(20, yPos);
            btnOpenWordReplacements.Width = 200;
            panel.Controls.Add(btnOpenWordReplacements);
            yPos += 40;
            
            btnOpenMultiActions = CreateButton("Edit Multi-Actions", (s, e) => OpenConfigFile("multi_actions.json"));
            btnOpenMultiActions.Location = new Point(20, yPos);
            btnOpenMultiActions.Width = 200;
            panel.Controls.Add(btnOpenMultiActions);
            yPos += 40;
            
            btnOpenEmojiMappings = CreateButton("Edit Emoji Mappings", (s, e) => OpenConfigFile("emoji_mappings.json"));
            btnOpenEmojiMappings.Location = new Point(20, yPos);
            btnOpenEmojiMappings.Width = 200;
            panel.Controls.Add(btnOpenEmojiMappings);
            yPos += 40;
            
            btnOpenVSCommands = CreateButton("Edit VS Commands", (s, e) => OpenConfigFile("vs_commands.json"));
            btnOpenVSCommands.Location = new Point(20, yPos);
            btnOpenVSCommands.Width = 200;
            panel.Controls.Add(btnOpenVSCommands);
            yPos += 50;
            
            btnReloadConfigs = CreateButton("Reload All Configs", (s, e) => ReloadConfigs());
            btnReloadConfigs.Location = new Point(20, yPos);
            btnReloadConfigs.Width = 200;
            panel.Controls.Add(btnReloadConfigs);
            yPos += 40;
            
            btnOpenSettingsFolder = CreateButton("Open Settings Folder", (s, e) => OpenSettingsFolder());
            btnOpenSettingsFolder.Location = new Point(20, yPos);
            btnOpenSettingsFolder.Width = 200;
            panel.Controls.Add(btnOpenSettingsFolder);
            
            tab.Controls.Add(panel);
            return tab;
        }

        private Panel CreateScrollablePanel()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(10),
                BackColor = DisplayMessage.SharedBackColor,
                ForeColor = DisplayMessage.SharedForeColor
            };
            return panel;
        }

        private Label CreateLabel(string text, ref int yPos, bool isHeader = false, float fontSize = 0)
        {
            var label = new Label
            {
                Text = text,
                Location = new Point(20, yPos),
                AutoSize = true,
                Font = isHeader 
                    ? new Font(DisplayMessage.SharedFont.FontFamily, fontSize > 0 ? fontSize : 10, FontStyle.Bold)
                    : fontSize > 0 
                        ? new Font(DisplayMessage.SharedFont.FontFamily, fontSize)
                        : DisplayMessage.SharedFont,
                ForeColor = DisplayMessage.SharedForeColor
            };
            yPos += isHeader ? 30 : 25;
            return label;
        }

        private CheckBox CreateCheckBox(string text, ref int yPos)
        {
            var checkBox = new CheckBox
            {
                Text = text,
                Location = new Point(20, yPos),
                AutoSize = true,
                Font = DisplayMessage.SharedFont,
                ForeColor = DisplayMessage.SharedForeColor,
                BackColor = DisplayMessage.SharedBackColor
            };
            yPos += 30;
            return checkBox;
        }

        private NumericUpDown CreateNumericUpDown(decimal min, decimal max, ref int yPos, bool allowDecimals = false, decimal increment = 1)
        {
            var numericUpDown = new NumericUpDown
            {
                Location = new Point(20, yPos),
                Width = 120,
                Minimum = min,
                Maximum = max,
                DecimalPlaces = allowDecimals ? 2 : 0,
                Increment = increment,
                Font = DisplayMessage.SharedFont,
                BackColor = Color.FromArgb(45, 45, 48),
                ForeColor = DisplayMessage.SharedForeColor
            };
            yPos += 35;
            return numericUpDown;
        }

        private TextBox CreateTextBox(ref int yPos)
        {
            var textBox = new TextBox
            {
                Location = new Point(20, yPos),
                Width = 380,
                Font = DisplayMessage.SharedFont,
                BackColor = Color.FromArgb(45, 45, 48),
                ForeColor = DisplayMessage.SharedForeColor
            };
            yPos += 35;
            return textBox;
        }

        private ComboBox CreateComboBox(string[] items, ref int yPos)
        {
            var comboBox = new ComboBox
            {
                Location = new Point(20, yPos),
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = DisplayMessage.SharedFont,
                BackColor = Color.FromArgb(45, 45, 48),
                ForeColor = DisplayMessage.SharedForeColor
            };
            comboBox.Items.AddRange(items);
            yPos += 35;
            return comboBox;
        }

        private ComboBox CreateFontComboBox(ref int yPos)
        {
            var comboBox = new ComboBox
            {
                Location = new Point(20, yPos),
                Width = 200,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = DisplayMessage.SharedFont,
                BackColor = Color.FromArgb(45, 45, 48),
                ForeColor = DisplayMessage.SharedForeColor
            };
            
            foreach (var family in FontFamily.Families)
            {
                comboBox.Items.Add(family.Name);
            }
            
            yPos += 35;
            return comboBox;
        }

        private void LoadSettings()
        {
            var settings = AppSettings.Instance;
            
            // General
            chkEnableNotifications.Checked = settings.Notifications.Enabled;
            numNotificationDuration.Value = settings.Notifications.DisplayDurationMs;
            chkDebugMode.Checked = settings.Behavior.DebugMode;
            chkShowDebugDialog.Checked = settings.Behavior.ShowDebugDialog;
            
            // Auto-Click
            numAutoClickDelay.Value = settings.AutoClick.DelayMs;
            chkAutoClickShowOverlay.Checked = settings.AutoClick.ShowOverlay;
            
            // Mouse Movement
            numMouseDefaultSpeed.Value = settings.MouseMovement.DefaultSpeed;
            numMouseMinSpeed.Value = settings.MouseMovement.MinSpeed;
            numMouseMaxSpeed.Value = settings.MouseMovement.MaxSpeed;
            numMouseSpeedStep.Value = settings.MouseMovement.SpeedStep;
            numMouseUpdateInterval.Value = settings.MouseMovement.UpdateIntervalMs;
            
            // Voice Dictation
            numVoiceAutoSubmitTimeout.Value = settings.VoiceDictation.AutoSubmitTimeoutMs;
            chkVoiceAutoDetectStop.Checked = settings.VoiceDictation.AutoDetectStop;
            numVoiceStopDebounce.Value = settings.VoiceDictation.StopDebounceMs;
            numVoiceFormWidth.Value = settings.VoiceDictation.FormWidth;
            numVoiceFormHeight.Value = settings.VoiceDictation.FormHeight;
            chkVoiceStartTransparent.Checked = settings.VoiceDictation.StartTransparent;
            numVoiceMarqueeInterval.Value = settings.VoiceDictation.MarqueeIntervalMs;
            
            // Appearance
            btnBackgroundColor.BackColor = ColorTranslator.FromHtml(settings.Appearance.BackgroundColor);
            btnForegroundColor.BackColor = ColorTranslator.FromHtml(settings.Appearance.ForegroundColor);
            cmbFontFamily.Text = settings.Appearance.FontFamily;
            numFontSize.Value = (decimal)settings.Appearance.FontSize;
            cmbRichTextFont.Text = settings.Appearance.RichTextFontFamily;
            numRichTextFontSize.Value = (decimal)settings.Appearance.RichTextFontSize;
            numAccessibilityMultiplier.Value = (decimal)settings.Appearance.AccessibilityFontMultiplier;
            
            // AI
            UpdateApiKeyStatus();
            txtModelName.Text = settings.AI.ModelName;
            chkEnableAIFallback.Checked = settings.AI.EnableFallback;
            txtPromptFile.Text = settings.AI.PromptFile;
            
            // Advanced
            chkLoggingEnabled.Checked = settings.Logging.Enabled;
            cmbLogLevel.Text = settings.Logging.LogLevel;
            chkClearLogOnStartup.Checked = settings.Logging.ClearOnStartup;
            numMaxLogSizeMB.Value = settings.Logging.MaxLogSizeMB;
            numMultiActionDelay.Value = settings.Behavior.MultiActionDelayMs;
            chkContinueOnError.Checked = settings.Behavior.ContinueOnError;
            numMessageBoxTimeout.Value = settings.Behavior.MessageBoxTimeoutMs;
            numVSCommandChordDelay.Value = settings.Behavior.VSCommandChordDelayMs;
        }

        private void SaveSettings()
        {
            var settings = AppSettings.Instance;
            
            // General
            settings.Notifications.Enabled = chkEnableNotifications.Checked;
            settings.Notifications.DisplayDurationMs = (int)numNotificationDuration.Value;
            settings.Behavior.DebugMode = chkDebugMode.Checked;
            settings.Behavior.ShowDebugDialog = chkShowDebugDialog.Checked;
            
            // Auto-Click
            settings.AutoClick.DelayMs = (int)numAutoClickDelay.Value;
            settings.AutoClick.ShowOverlay = chkAutoClickShowOverlay.Checked;
            
            // Mouse Movement
            settings.MouseMovement.DefaultSpeed = (int)numMouseDefaultSpeed.Value;
            settings.MouseMovement.MinSpeed = (int)numMouseMinSpeed.Value;
            settings.MouseMovement.MaxSpeed = (int)numMouseMaxSpeed.Value;
            settings.MouseMovement.SpeedStep = (int)numMouseSpeedStep.Value;
            settings.MouseMovement.UpdateIntervalMs = (int)numMouseUpdateInterval.Value;
            
            // Voice Dictation
            settings.VoiceDictation.AutoSubmitTimeoutMs = (int)numVoiceAutoSubmitTimeout.Value;
            settings.VoiceDictation.AutoDetectStop = chkVoiceAutoDetectStop.Checked;
            settings.VoiceDictation.StopDebounceMs = (int)numVoiceStopDebounce.Value;
            settings.VoiceDictation.FormWidth = (int)numVoiceFormWidth.Value;
            settings.VoiceDictation.FormHeight = (int)numVoiceFormHeight.Value;
            settings.VoiceDictation.StartTransparent = chkVoiceStartTransparent.Checked;
            settings.VoiceDictation.MarqueeIntervalMs = (int)numVoiceMarqueeInterval.Value;
            
            // Appearance
            settings.Appearance.BackgroundColor = ColorTranslator.ToHtml(btnBackgroundColor.BackColor);
            settings.Appearance.ForegroundColor = ColorTranslator.ToHtml(btnForegroundColor.BackColor);
            settings.Appearance.FontFamily = cmbFontFamily.Text;
            settings.Appearance.FontSize = (float)numFontSize.Value;
            settings.Appearance.RichTextFontFamily = cmbRichTextFont.Text;
            settings.Appearance.RichTextFontSize = (float)numRichTextFontSize.Value;
            settings.Appearance.AccessibilityFontMultiplier = (float)numAccessibilityMultiplier.Value;
            
            // AI
            txtModelName.Text = settings.AI.ModelName;
            settings.AI.EnableFallback = chkEnableAIFallback.Checked;
            settings.AI.PromptFile = txtPromptFile.Text;
            
            // Advanced
            settings.Logging.Enabled = chkLoggingEnabled.Checked;
            settings.Logging.LogLevel = cmbLogLevel.Text;
            settings.Logging.ClearOnStartup = chkClearLogOnStartup.Checked;
            settings.Logging.MaxLogSizeMB = (int)numMaxLogSizeMB.Value;
            settings.Behavior.MultiActionDelayMs = (int)numMultiActionDelay.Value;
            settings.Behavior.ContinueOnError = chkContinueOnError.Checked;
            settings.Behavior.MessageBoxTimeoutMs = (int)numMessageBoxTimeout.Value;
            settings.Behavior.VSCommandChordDelayMs = (int)numVSCommandChordDelay.Value;
            
            try
            {
                settings.Save();
                MessageBox.Show("Settings saved successfully!", "Success", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                Logger.LogInfo("Settings saved successfully");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving settings: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                Logger.LogError($"Error saving settings: {ex.Message}");
            }
        }

        private void BtnSave_Click(object? sender, EventArgs e)
        {
            SaveSettings();
            DialogResult = DialogResult.OK;
            Close();
        }

        private void BtnApply_Click(object? sender, EventArgs e)
        {
            SaveSettings();
        }

        private void PickColor(ref Button button)
        {
            using (var colorDialog = new ColorDialog())
            {
                colorDialog.Color = button.BackColor;
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    button.BackColor = colorDialog.Color;
                }
            }
        }

        private void UpdateApiKeyStatus()
        {
            if (AppSettings.Instance.AI.IsConfigured)
            {
                lblApiKeyStatus.Text = "API Key: Configured ✓";
                lblApiKeyStatus.ForeColor = Color.LightGreen;
            }
            else
            {
                lblApiKeyStatus.Text = "API Key: Not configured";
                lblApiKeyStatus.ForeColor = Color.Orange;
            }
        }

        private void ConfigureApiKey()
        {
            MessageBox.Show(
                "To configure the OpenAI API key, set the OPENAI_API_KEY environment variable.\n\n" +
                "You can do this:\n" +
                "1. System-wide: Settings → System → About → Advanced system settings → Environment Variables\n" +
                "2. User-level: Same as above, but in the 'User variables' section\n" +
                "3. Restart NaturalCommands after setting the variable",
                "Configure API Key",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        private void BrowsePromptFile()
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Markdown files (*.md)|*.md|Text files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // Make path relative to app directory
                    string appDir = AppDomain.CurrentDomain.BaseDirectory;
                    string relativePath = openFileDialog.FileName;
                    if (relativePath.StartsWith(appDir))
                    {
                        relativePath = relativePath.Substring(appDir.Length);
                    }
                    txtPromptFile.Text = relativePath;
                }
            }
        }

        private void OpenConfigFile(string fileName)
        {
            try
            {
                string filePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName);
                if (System.IO.File.Exists(filePath))
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = filePath,
                        UseShellExecute = true
                    });
                }
                else
                {
                    MessageBox.Show($"File not found: {fileName}", "Error", 
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening file: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ReloadConfigs()
        {
            try
            {
                AppSettings.Reload();
                LoadSettings();
                MessageBox.Show("Configuration reloaded successfully!", "Success", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                Logger.LogInfo("Configuration reloaded from settings form");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reloading configuration: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void OpenSettingsFolder()
        {
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = AppDomain.CurrentDomain.BaseDirectory,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening folder: {ex.Message}", "Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
