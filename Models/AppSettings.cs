using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace NaturalCommands.Models
{
    /// <summary>
    /// Comprehensive application settings that can be configured via the Settings form.
    /// Settings are persisted to settings.json in the application directory.
    /// </summary>
    public class AppSettings
    {
        private static AppSettings? _instance;
        private static readonly object _lock = new object();
        private static readonly string SettingsPath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory, "settings.json");

        // Auto-Click Settings
        public AutoClickSettings AutoClick { get; set; } = new AutoClickSettings();

        // Mouse Movement Settings
        public MouseMovementSettings MouseMovement { get; set; } = new MouseMovementSettings();

        // Voice Dictation Settings
        public VoiceDictationSettings VoiceDictation { get; set; } = new VoiceDictationSettings();

        // Appearance Settings
        public AppearanceSettings Appearance { get; set; } = new AppearanceSettings();

        // AI Integration Settings
        public AISettings AI { get; set; } = new AISettings();

        // Logging Settings
        public LoggingSettings Logging { get; set; } = new LoggingSettings();

        // Notification Settings
        public NotificationSettings Notifications { get; set; } = new NotificationSettings();

        // Behavior Settings
        public BehaviorSettings Behavior { get; set; } = new BehaviorSettings();

        // Hotkey Settings
        public HotkeySettings Hotkeys { get; set; } = new HotkeySettings();

        /// <summary>
        /// Gets the singleton instance of AppSettings.
        /// </summary>
        public static AppSettings Instance
        {
            get
            {
                lock (_lock)
                {
                    if (_instance == null)
                    {
                        _instance = Load();
                    }
                    return _instance;
                }
            }
        }

        /// <summary>
        /// Loads settings from settings.json or creates default settings if file doesn't exist.
        /// </summary>
        public static AppSettings Load()
        {
            try
            {
                if (File.Exists(SettingsPath))
                {
                    var json = File.ReadAllText(SettingsPath);
                    var settings = JsonSerializer.Deserialize<AppSettings>(json, new JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true,
                        ReadCommentHandling = JsonCommentHandling.Skip,
                        WriteIndented = true
                    });
                    return settings ?? new AppSettings();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading settings: {ex.Message}");
            }

            // Return default settings
            return new AppSettings();
        }

        /// <summary>
        /// Saves current settings to settings.json.
        /// </summary>
        public void Save()
        {
            try
            {
                var json = JsonSerializer.Serialize(this, new JsonSerializerOptions
                {
                    WriteIndented = true,
                    DefaultIgnoreCondition = JsonIgnoreCondition.Never
                });
                File.WriteAllText(SettingsPath, json);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Failed to save settings: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Reloads settings from disk.
        /// </summary>
        public static void Reload()
        {
            lock (_lock)
            {
                _instance = Load();
            }
        }
    }

    public class AutoClickSettings
    {
        /// <summary>
        /// Delay in milliseconds before auto-click triggers (100-5000ms).
        /// </summary>
        public int DelayMs { get; set; } = 2000;

        /// <summary>
        /// Whether auto-click mode is currently enabled.
        /// </summary>
        public bool Enabled { get; set; } = false;

        /// <summary>
        /// Whether to show the countdown overlay.
        /// </summary>
        public bool ShowOverlay { get; set; } = true;
    }

    public class MouseMovementSettings
    {
        /// <summary>
        /// Default mouse movement speed in pixels per tick (2-50).
        /// </summary>
        public int DefaultSpeed { get; set; } = 5;

        /// <summary>
        /// Minimum allowed speed.
        /// </summary>
        public int MinSpeed { get; set; } = 2;

        /// <summary>
        /// Maximum allowed speed.
        /// </summary>
        public int MaxSpeed { get; set; } = 50;

        /// <summary>
        /// Speed adjustment step when using faster/slower commands.
        /// </summary>
        public int SpeedStep { get; set; } = 5;

        /// <summary>
        /// Movement update interval in milliseconds (controls smoothness).
        /// </summary>
        public int UpdateIntervalMs { get; set; } = 16; // ~60 FPS
    }

    public class VoiceDictationSettings
    {
        /// <summary>
        /// Auto-submit timeout in milliseconds (0 = disabled).
        /// </summary>
        public int AutoSubmitTimeoutMs { get; set; } = 0;

        /// <summary>
        /// Enable automatic detection of dictation stop.
        /// </summary>
        public bool AutoDetectStop { get; set; } = false;

        /// <summary>
        /// Dictation stop debounce delay in milliseconds.
        /// </summary>
        public int StopDebounceMs { get; set; } = 850;

        /// <summary>
        /// Default form width.
        /// </summary>
        public int FormWidth { get; set; } = 1200;

        /// <summary>
        /// Default form height.
        /// </summary>
        public int FormHeight { get; set; } = 800;

        /// <summary>
        /// Whether form should start in transparent mode.
        /// </summary>
        public bool StartTransparent { get; set; } = false;

        /// <summary>
        /// Marquee scroll speed (timer interval in ms).
        /// </summary>
        public int MarqueeIntervalMs { get; set; } = 28;
    }

    public class AppearanceSettings
    {
        /// <summary>
        /// Background color in hex format (e.g., "#1E1E1E").
        /// </summary>
        public string BackgroundColor { get; set; } = "#1E1E1E";

        /// <summary>
        /// Foreground/text color in hex format.
        /// </summary>
        public string ForegroundColor { get; set; } = "#FFFFFF";

        /// <summary>
        /// Default font family name.
        /// </summary>
        public string FontFamily { get; set; } = "Segoe UI";

        /// <summary>
        /// Default font size in points.
        /// </summary>
        public float FontSize { get; set; } = 9.75f;

        /// <summary>
        /// Rich text box font family.
        /// </summary>
        public string RichTextFontFamily { get; set; } = "Cascadia Code";

        /// <summary>
        /// Rich text box font size.
        /// </summary>
        public float RichTextFontSize { get; set; } = 15.75f;

        /// <summary>
        /// Font size multiplier for accessibility.
        /// </summary>
        public float AccessibilityFontMultiplier { get; set; } = 1.4f;
    }

    public class AISettings
    {
        /// <summary>
        /// OpenAI API key (stored in environment variable for security).
        /// This field is not serialized - use environment variable OPENAI_API_KEY.
        /// </summary>
        [JsonIgnore]
        public string ApiKey => Environment.GetEnvironmentVariable("OPENAI_API_KEY") ?? "";

        /// <summary>
        /// Whether the API key is configured.
        /// </summary>
        [JsonIgnore]
        public bool IsConfigured => !string.IsNullOrEmpty(ApiKey);

        /// <summary>
        /// AI model name to use.
        /// </summary>
        public string ModelName { get; set; } = "gpt-4.1";

        /// <summary>
        /// Enable AI fallback for command interpretation.
        /// </summary>
        public bool EnableFallback { get; set; } = true;

        /// <summary>
        /// Custom prompt file path (relative to app directory).
        /// </summary>
        public string PromptFile { get; set; } = "openai_prompt.md";
    }

    public class LoggingSettings
    {
        /// <summary>
        /// Enable logging to file.
        /// </summary>
        public bool Enabled { get; set; } = true;

        /// <summary>
        /// Log level: Debug, Info, Warning, Error.
        /// </summary>
        public string LogLevel { get; set; } = "Info";

        /// <summary>
        /// Clear log file on application startup.
        /// </summary>
        public bool ClearOnStartup { get; set; } = true;

        /// <summary>
        /// Custom log file path (relative to app directory, empty = default).
        /// </summary>
        public string CustomLogPath { get; set; } = "";

        /// <summary>
        /// Maximum log file size in MB before rotation (0 = unlimited).
        /// </summary>
        public int MaxLogSizeMB { get; set; } = 10;
    }

    public class NotificationSettings
    {
        /// <summary>
        /// Display duration for balloon tip notifications in milliseconds.
        /// </summary>
        public int DisplayDurationMs { get; set; } = 5000;

        /// <summary>
        /// Enable balloon tip notifications.
        /// </summary>
        public bool Enabled { get; set; } = true;

        /// <summary>
        /// Maximum tooltip text length.
        /// </summary>
        public int MaxTooltipLength { get; set; } = 60;
    }

    public class BehaviorSettings
    {
        /// <summary>
        /// Default delay between multi-actions in milliseconds.
        /// </summary>
        public int MultiActionDelayMs { get; set; } = 250;

        /// <summary>
        /// Continue executing multi-actions on error.
        /// </summary>
        public bool ContinueOnError { get; set; } = true;

        /// <summary>
        /// Message box auto-close timeout in milliseconds.
        /// </summary>
        public int MessageBoxTimeoutMs { get; set; } = 5000;

        /// <summary>
        /// Visual Studio command chord delay in milliseconds.
        /// </summary>
        public int VSCommandChordDelayMs { get; set; } = 80;

        /// <summary>
        /// Enable debug mode.
        /// </summary>
        public bool DebugMode { get; set; } = false;

        /// <summary>
        /// Show debug input dialog on startup (debug mode only).
        /// </summary>
        public bool ShowDebugDialog { get; set; } = false;
    }

    public class HotkeySettings
    {
        /// <summary>
        /// Voice dictation hotkey key code (default: H).
        /// </summary>
        public int VoiceDictationKey { get; set; } = 0x48; // H key

        /// <summary>
        /// Voice dictation hotkey modifiers (1=Alt, 2=Ctrl, 4=Shift, 8=Win).
        /// Default: Ctrl+Win = 10.
        /// </summary>
        public int VoiceDictationModifiers { get; set; } = 10; // Ctrl+Win

        /// <summary>
        /// Enable voice dictation hotkey registration.
        /// </summary>
        public bool EnableVoiceDictationHotkey { get; set; } = true;
    }
}
