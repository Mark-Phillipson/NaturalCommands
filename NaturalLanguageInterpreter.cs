using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using WindowsInput;
using WindowsInput.Native;


using OpenAI;
using OpenAI.Chat;
using OpenAI.Models;
using NaturalCommands;
using NaturalCommands.Models;

namespace NaturalCommands
{
    // Win32 API imports and constants for window style and class name
    internal static class Win32Api
    {
        public const int GWL_STYLE = -16;
        public const int WS_MAXIMIZEBOX = 0x00010000;

        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true, CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        public static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);
    }
    // IMPORTANT: This class is already too large. Do NOT add new methods here.
    // Any new functionality should be implemented in a new class and referenced as needed.
    // Please refactor existing logic into smaller, focused classes where possible.
    public class NaturalLanguageInterpreter
    // Action type for Visual Studio command execution
    {
        /// <summary>
        /// Checks if Visual Studio is the active window.
        /// </summary>
        public static bool IsVisualStudioActive()
        {
            var procName = NaturalCommands.CurrentApplicationHelper.GetCurrentProcessName();
            return procName == "devenv";
        }
        /// <summary>
        /// Ensures the directory for the log file exists.
        /// </summary>
        // Expanded app mapping for natural language launching
        private static readonly Dictionary<string, string> AppMappings = new(StringComparer.OrdinalIgnoreCase)
        {
            { "calculator", "calc.exe" },
            { "calc", "calc.exe" },
            { "notepad", "notepad.exe" },
            { "edge", "msedge.exe" },
            { "microsoft edge", "msedge.exe" },
            { "chrome", "chrome.exe" },
            { "code", "code.exe" },
            { "visual studio", "devenv.exe" },
            { "outlook", "outlook.exe" },
            { "explorer", "explorer.exe" },
            { "word", "winword.exe" },
            { "excel", "excel.exe" },
            { "powerpoint", "powerpnt.exe" },
            { "teams", "Teams.exe" },
            { "onenote", "onenote.exe" },
            { "paint", "mspaint.exe" },
            { "microsoft paint", "mspaint.exe" },
            { "terminal", "wt.exe" },
            { "windows terminal", "wt.exe" },
            { "cmd", "wt.exe" }, // Always prefer Windows Terminal
            { "command prompt", "wt.exe" },
            { "steam", "steam://open/main" },
            { "skype", "skype.exe" },
            { "zoom", "zoom.exe" },
            { "slack", "slack.exe" }
        };
        /// <summary>
        /// Uses OpenAI API to interpret text and return an ActionBase (AI fallback).
        /// </summary>
        public async System.Threading.Tasks.Task<ActionBase?> InterpretWithAIAsync(string text)
        {
            // Read API key from environment variable
            string? apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
            if (string.IsNullOrWhiteSpace(apiKey))
            {
                NaturalCommands.Helpers.Logger.LogError("OPENAI_API_KEY environment variable not set.");
                return null;
            }
            // Set default model name
            string modelName = "gpt-4.1";

            // Read prompt from markdown file
            string promptPath = "openai_prompt.md";
            string prompt;
            try
            {
                prompt = File.ReadAllText(promptPath);
            }
            catch (Exception ex)
            {
                NaturalCommands.Helpers.Logger.LogError($"Failed to read {promptPath}: {ex.Message} Using default prompt.");
                prompt = "You are an assistant that interprets natural language commands for Windows automation. Output a JSON object for the closest matching action.";
            }

            NaturalCommands.Helpers.Logger.LogDebug($"[AI] Fallback triggered for: {text}");
            // Write the latest prompt to a separate file (overwrites previous file).
            // This is intentionally NOT appended to the normal log file.
            WriteLatestPromptFile(prompt, text);
            // Do NOT log the prompt anymore
            try
            {
                var chatClient = new ChatClient(modelName, apiKey);
                var messages = new List<ChatMessage>
                {
                    new SystemChatMessage(prompt),
                    new UserChatMessage(text)
                };
                var completionResult = await chatClient.CompleteChatAsync(messages);
                var completion = completionResult.Value;
                var message = completion.Content[0].Text;
                NaturalCommands.Helpers.Logger.LogDebug($"[AI] Raw response: {message}");
                if (!string.IsNullOrWhiteSpace(message))
                {
                    try
                    {
                        var json = System.Text.Json.JsonDocument.Parse(message);
                        var root = json.RootElement;
                        if (root.TryGetProperty("type", out var typeProp))
                        {
                            string type = typeProp.GetString() ?? "";
                            switch (type)
                            {
                                case "MoveWindowAction":
                                    return new MoveWindowAction(
                                        root.GetProperty("Target").GetString() ?? "active",
                                        root.GetProperty("Monitor").GetString() ?? "current",
                                        root.GetProperty("Position").GetString(),
                                        root.TryGetProperty("WidthPercent", out var wp) ? wp.GetInt32() : (int?)null,
                                        root.TryGetProperty("HeightPercent", out var hp) ? hp.GetInt32() : (int?)null
                                    );
                                case "LaunchAppAction":
                                    // Support both "AppExe" and "AppIdOrPath" field names
                                    string? appExe = null;
                                    if (root.TryGetProperty("AppExe", out var appExeProp))
                                        appExe = appExeProp.GetString();
                                    else if (root.TryGetProperty("AppIdOrPath", out var appIdProp))
                                        appExe = appIdProp.GetString();

                                    return new LaunchAppAction(appExe ?? "");
                                case "SendKeysAction":
                                    return new SendKeysAction(root.GetProperty("KeysText").GetString() ?? "");
                                case "OpenFolderAction":
                                    return new OpenFolderAction(root.GetProperty("KnownFolder").GetString() ?? "");
                                // "SetWindowAlwaysOnTopAction" intentionally not supported: feature disabled.
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        NaturalCommands.Helpers.Logger.LogError($"Failed to parse OpenAI response: {ex.Message}\nResponse: {message}");
                    }
                }
            }
            catch (Exception ex)
            {
                NaturalCommands.Helpers.Logger.LogError($"[AI] OpenAI API call failed: {ex.Message}");
            }
            return null;
        }
        // Helper to remove polite modifiers from input
        private static string RemovePoliteModifiers(string text)
        {
            var politeWords = new[] { "please", "could you", "would you", "can you", "may you", "kindly", "will you", "would you kindly" };
            foreach (var word in politeWords)
            {
                text = text.Replace(word, "", StringComparison.InvariantCultureIgnoreCase);
            }
            return text.Trim();
        }

        // Word replacement functionality moved to `WordReplacementLoader` helper class.


        /// <summary>
        /// Writes the latest AI prompt (system prompt + user input) to a separate file.
        /// Overwrites any previous content so only the latest prompt is kept.
        /// The file is intentionally separate from the normal `app.log`.
        /// </summary>
        private static void WriteLatestPromptFile(string systemPrompt, string userText)
        {
            try
            {
                string latestPath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "bin", "latest_ai_prompt.md"));
                var dir = Path.GetDirectoryName(latestPath);
                if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                    Directory.CreateDirectory(dir);

                var content = $"# System Prompt\n\n{systemPrompt}\n\n# User Input\n\n{userText}\n";
                File.WriteAllText(latestPath, content);
            }
            catch (Exception ex)
            {
                try { NaturalCommands.Helpers.Logger.LogError($"Failed to write latest AI prompt file: {ex.Message}"); } catch { }
            }
        }
        // Central list of available commands/actions for AI matching
        public static readonly List<(string Command, string Description)> AvailableCommands = new()
        {
            ("maximize window", "Maximize the active window"),
            ("move window to left half", "Move the active window to the left half of the screen"),
            ("move window to right half", "Move the active window to the right half of the screen"),
            ("move window to other monitor", "Move the active window to the next monitor"),
            ("open downloads", "Open the Downloads folder"),
            ("open documents", "Open the Documents folder"),
            ("open settings", "Open the application settings"),
            ("close tab", "Close the current tab in supported applications"),
            ("send keys", "Send a key sequence to the active window"),
            ("launch app", "Launch a specified application"),
            ("focus app", "Focus a specified application window"),
            ("focus window <name>", "Focus a window by its name (e.g. focus window Zoom)"),
            ("focus <window name>", "Focus a window by its name (e.g. focus Zoom)"),
            ("show help", "Show help and available commands"),
            ("natural dictate", "Open the voice dictation form (speak or type natural language commands)"),
            ("show letters", "Display letter labels on clickable UI elements for voice-based navigation"),
            ("identify <target>", "Identify a visual target and click when confident, otherwise show numbered options"),
            ("id <target>", "Shorthand for identify: identify a visual target (e.g. id telegram)"),
            ("show candidates", "Show numbered candidates from the latest visual identify command"),
            ("choose <number>", "Choose a numbered visual target candidate"),
            ("emoji set <name> <emoji>", "Set an emoji for a named shortcut (e.g. emoji set happy 😀)"),
            ("emoji <name>", "Insert the configured emoji for the given name"),
            ("emoji <emoji>", "Insert the given emoji immediately"),
            ("move <direction>", "Start moving the mouse continuously (e.g. move up, move left, move down right, mouse move left)"),
            ("mouse move <direction>", "Start moving the mouse continuously (e.g. mouse move up, mouse left)"),
            ("stop mouse", "Stop mouse movement"),
            ("stop click", "Stop mouse movement and perform a left click"),
            ("stop right click", "Stop mouse movement and perform a right click"),
            ("mouse stop", "Stop mouse movement"),
            ("faster", "Increase mouse movement speed"),
            ("slower", "Decrease mouse movement speed"),
            ("mouse faster", "Increase mouse movement speed"),
            ("mouse slower", "Decrease mouse movement speed"),
            ("auto click", "Enable auto-click when mouse is idle (default 2000ms delay)"),
            ("enable auto click", "Enable auto-click mode"),
            ("start auto click", "Enable auto-click mode"),
            ("stop auto click", "Disable auto-click mode"),
            ("disable auto click", "Disable auto-click mode"),
            ("auto click off", "Disable auto-click mode"),
            ("auto click faster", "Increase auto-click speed (shorten delay)"),
            ("auto click speed up", "Increase auto-click speed (shorten delay)"),
            ("auto click speedup", "Increase auto-click speed (shorten delay)"),
            ("auto click slower", "Decrease auto-click speed (lengthen delay)"),
            ("auto click slow down", "Decrease auto-click speed (lengthen delay)"),
            ("auto click slowdown", "Decrease auto-click speed (lengthen delay)")
        };

        // Optional emoji mapping for commands. Map a command phrase to a small emoji
        // string that will be displayed next to the command in the 'what can I say' UI.
        // Emoji logic moved to EmojiManager.cs. Use EmojiManager.SetCommandEmoji, EmojiManager.GetCommandEmoji, EmojiManager.GetAllEmojiMappings instead.

        // File used to persist emoji mappings so they can be added over time.
        // Will look for a file next to the executable (copied by the build as content).
        private static readonly string EmojiMappingsPath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "emoji_mappings.json"));

        // Back-compat helper used throughout the file: delegate to centralized Logger
        private static void AppendLog(string message)
        {
            try { NaturalCommands.Helpers.Logger.Log(message.TrimEnd()); } catch { }
        }

        // Static ctor: load persisted mappings (if any) on first access
        static NaturalLanguageInterpreter()
        {
            // Emoji mappings now loaded by EmojiManager
            // Load optional word replacements (e.g., 'closed' -> 'close') to make parsing deterministic
            WordReplacementLoader.Load();
            // Load user-defined multi-action commands (multi_actions.json)
            try { NaturalCommands.Helpers.MultiActionLoader.Load(); } catch { }
            // Load Talon phrase catalog for fallback matching
            try { NaturalCommands.Helpers.TalonCommandCatalog.EnsureLoaded(); } catch { }
        }

        // Emoji mapping API now provided by EmojiManager.cs

        // Visual Studio specific commands
        public static readonly List<(string Command, string Description)> VisualStudioCommands = new()
            {
                ("build the solution", "Build the entire solution"),
                ("build the project", "Build the current project"),
                ("start debugging", "Start debugging the startup project"),
                ("start application", "Start without debugging"),
                ("stop debugging", "Stop debugging"),
                ("close tab", "Close the current document tab"),
                ("format document", "Format the current document"),
                ("find in files", "Open the Find in Files dialog"),
                ("go to definition", "Go to definition of symbol"),
                ("rename symbol", "Rename the selected symbol"),
                ("show solution explorer", "Focus Solution Explorer"),
                ("open recent files", "Show recent files"),
            };

        // Explicit canonical mappings from natural phrases to Visual Studio command canonical names.
        // These are preferred over fuzzy-matching exported commands because exported lists may contain
        // context-menu entries that are not the canonical top-level commands (which caused failures).
        private static readonly Dictionary<string, string> VisualStudioCanonicalMappings = new(StringComparer.OrdinalIgnoreCase)
        {
            { "build the solution", "Build.BuildSolution" },
            { "build solution", "Build.BuildSolution" },
            { "build the project", "Build.BuildProject" },
            { "build project", "Build.BuildProject" },
            { "clean solution", "Build.CleanSolution" },
            { "clean the solution", "Build.CleanSolution" },
            { "start debugging", "Debug.Start" },
            { "start application", "Debug.StartWithoutDebugging" },
            { "stop debugging", "Debug.StopDebugging" },
            { "close tab", "Window.CloseDocumentWindow" },
            { "close tool window", "Window.CloseToolWindow" },
            { "close the tool window", "Window.CloseToolWindow" },
            { "close current tool window", "Window.CloseToolWindow" },
            { "format document", "Edit.FormatDocument" },
            { "find in files", "Edit.FindinFiles" },
            { "go to definition", "Edit.GoToDefinition" },
            { "rename symbol", "Refactor.Rename" },
            { "show solution explorer", "View.SolutionExplorer" },
            { "open recent files", "File.RecentFiles" }
        };

        // Mapping from common tool window captions (as spoken or seen in the UI) to their canonical Visual Studio command names.
        // This ensures that natural language like "error list", "output window", etc. will always focus the correct tool window,
        // and avoids accidental matches to context menu or non-window commands. This mapping is checked BEFORE any fuzzy or exported command matches.
        // To add support for a new tool window, simply add its caption (as spoken or as it appears in Visual Studio) and the corresponding View.* command here.
        private static readonly Dictionary<string, string> VisualStudioToolWindowMappings = new(StringComparer.OrdinalIgnoreCase)
        {
            { "error list", "View.ErrorList" },
            { "output window", "View.Output" },
            { "output", "View.Output" },
            { "solution explorer", "View.SolutionExplorer" },
            { "team explorer", "View.TeamExplorer" },
            { "task list", "View.TaskList" },
            { "properties window", "View.PropertiesWindow" },
            { "properties", "View.PropertiesWindow" },
            { "class view", "View.ClassView" },
            { "object browser", "View.ObjectBrowser" },
            { "call hierarchy", "View.CallHierarchy" },
            { "bookmark window", "View.BookmarkWindow" },
            { "bookmarks", "View.BookmarkWindow" },
            { "find results", "View.FindResults1" },
            { "find results 1", "View.FindResults1" },
            { "find results 2", "View.FindResults2" },
            { "pending changes", "View.PendingChanges" },
            { "git changes", "View.GitChanges" },
            { "git repository", "View.GitRepository" },
            { "diagnostic tools", "Debug.ShowDiagnosticTools" },
            { "immediate window", "Debug.Immediate" },
            { "immediate", "Debug.Immediate" },
            { "autos window", "Debug.Autos" },
            { "autos", "Debug.Autos" },
            { "locals window", "Debug.Locals" },
            { "locals", "Debug.Locals" },
            { "watch window", "Debug.Watch" },
            { "watch", "Debug.Watch" },
            { "call stack", "Debug.CallStack" },
            { "breakpoints", "Debug.Breakpoints" },
            { "exception settings", "Debug.ExceptionSettings" },
            { "test explorer", "TestExplorer.ShowTestExplorer" },
            { "test window", "TestExplorer.ShowTestExplorer" },
            { "live unit testing window", "TestExplorer.ShowLiveUnitTestingWindow" },
            { "live unit testing", "TestExplorer.ShowLiveUnitTestingWindow" },
            { "solution explorer window", "View.SolutionExplorer" },
            { "output pane", "View.Output" },
            { "task pane", "View.TaskList" },
            { "error pane", "View.ErrorList" },
            { "explorer", "View.SolutionExplorer" },
            { "search results", "View.FindResults1" },
            { "search results 1", "View.FindResults1" },
            { "search results 2", "View.FindResults2" },
            { "pending changes window", "View.PendingChanges" },
            { "git window", "View.GitChanges" },
            { "repository window", "View.GitRepository" },
            { "diagnostics", "Debug.ShowDiagnosticTools" },
            { "breakpoint window", "Debug.Breakpoints" },
            { "exception window", "Debug.ExceptionSettings" },
            { "test", "TestExplorer.ShowTestExplorer" },
            { "tests", "TestExplorer.ShowTestExplorer" }
        };

        // Popular commands that override any matches
        private static readonly Dictionary<string, ActionBase> PopularCommands = new(StringComparer.OrdinalIgnoreCase)
        {
            { "debug application", new ExecuteVSCommandAction("Debug.Start") },
            { "run application", new ExecuteVSCommandAction("Debug.StartWithoutDebugging") },
            { "stop application", new ExecuteVSCommandAction("Debug.StopDebugging") },
            // Ensure 'focus' always triggers Ctrl+Alt+Tab
            { "focus", new SendKeysAction("ctrl alt tab") },
            // Voice dictation trigger: opens the voice dictation form (auto-submit 5s)
            // Use TimeoutMs=0 so the dictation form does not auto-submit — waits for manual Submit
            { "dictate", new OpenVoiceDictationFormAction(0) }
        };

        // VS Code specific commands
        public static readonly List<(string Command, string Description)> VSCodeCommands = new()
            {
                ("open file", "Open a file"),
                ("open folder", "Open a folder"),
                ("close tab", "Close the current tab"),
                ("format document", "Format the current document"),
                ("find in files", "Find in files"),
                ("go to definition", "Go to definition of symbol"),
                ("rename symbol", "Rename the selected symbol"),
                ("show explorer", "Show Explorer"),
                ("show source control", "Show Source Control"),
                ("show extensions", "Show Extensions"),
                ("start debugging", "Start debugging"),
                ("stop debugging", "Stop debugging"),
            };

        // Enhanced 'what can I say' logic
        public static void ShowAvailableCommands()
        {
            // If this method shows a dialog for long lists, we set this flag so
            // the subsequent ShowHelpAction execution can skip the redundant tray notification.
            // This avoids showing the same information twice (dialog + balloon).
            // It is reset after the ShowHelpAction is processed.
            // Note: internal flag; not exposed publicly.
            _suppressNextHelpNotification = false;
            string? procName = NaturalCommands.CurrentApplicationHelper.GetCurrentProcessName();

            if (procName == "devenv")
            {
                try
                {
                    var window = new SearchVisualStudioCommandsWPF();
                    window.ShowDialog();
                    return;
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show($"Error opening search window: {ex.Message}");
                }
            }

            List<(string Command, string Description)> commands;
            string appLabel;
            if (procName == "devenv")
            {
                commands = VisualStudioCommands;
                appLabel = "Visual Studio";
            }
            else if (procName == "code")
            {
                commands = VSCodeCommands;
                appLabel = "VS Code";
            }
            else if (procName == "windowsterminal" || procName == "WindowsTerminal")
            {
                // Combine Windows Terminal shortcuts with dotnet CLI commands
                commands = new List<(string Command, string Description)>(CommandDefinitions.WindowsTerminalCommands);
                // Add a section header for dotnet commands
                commands.Add(("--- .NET CLI Commands ---", ""));
                commands.AddRange(CommandDefinitions.DotNetCommands.Select(c => (c.Command, c.Description)));
                appLabel = "Windows Terminal";
            }
            else
            {
                commands = AvailableCommands;
                appLabel = "General";
            }

            // Format command list for display. If an emoji is configured for the command,
            // show it before the command text (e.g. "📥 open downloads: Open the Downloads folder").
            var lines = commands.Select(c =>
            {
                var emoji = EmojiManager.GetCommandEmoji(c.Command);
                if (!string.IsNullOrEmpty(emoji))
                    return $"- {emoji} {c.Command}: {c.Description}";
                return $"- {c.Command}: {c.Description}";
            }).ToList();
            lines.Add("- refresh Visual Studio shortcuts: Reload the latest keyboard shortcuts from Visual Studio settings");
            string message = $"Available commands:\n\n" + string.Join("\n", lines);

            // If command list is long, show in dialog and use notification as pointer
            if (lines.Count > 8)
            {
                try
                {
                    var form = new DictationBoxMSP.AvailableCommandsForm();
                    form.Text = "Available Commands";
                    // Show modal so callers don't continue until user closes the list.
                    form.ShowDialog();
                    _suppressNextHelpNotification = true;
                }
                catch (Exception)
                {
                    // Fallback to the original DisplayMessage if the new form fails to open
                    var dlg = new DictationBoxMSP.DisplayMessage(message, 60000, "Available Commands"); // 60 seconds, custom title
                    System.Windows.Forms.Application.Run(dlg); // Auto-close after timeout
                    _suppressNextHelpNotification = true;
                }
            }
            else
            {
                NaturalCommands.TrayNotificationHelper.ShowNotification($"{appLabel} Commands", string.Join("\n", lines), 7000);
            }
            // Also log to app.log for reference
            AppendLog($"[INFO] {appLabel} Supported Commands:\n{message}\n");
        }

        // Internal flag used to avoid showing a tray notification when the dialog
        // has already presented the available commands to the user.
        private static bool _suppressNextHelpNotification = false;
        // P/Invoke for MonitorFromWindow
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

        // P/Invoke for GetMonitorInfo
        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFOEX lpmi);

        // MONITORINFOEX struct
        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential, CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        public struct MONITORINFOEX
        {
            public int cbSize;
            public RECT rcMonitor;
            public RECT rcWork;
            public uint dwFlags;
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst = 32)]
            public string szDevice;
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
        // Action types are now defined in ActionModels.cs

        // P/Invoke for SetWindowPos
        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        // P/Invoke for ShowWindow
        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        // P/Invoke helpers for performing mouse clicks from interpreter
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetCursorPos(int X, int Y);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool GetCursorPos(out System.Drawing.Point lpPoint);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(System.Drawing.Point Point);

        // Test hooks: allow unit tests to inject candidate lists and intercept clicks.
        public static System.Func<string, System.Collections.Generic.List<NaturalCommands.Models.VisualTargetCandidate>>? VisualIdentifyCandidatesOverride;
        public static System.Action<System.Drawing.Point>? ClickOverride;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, int dwExtraInfo);

        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;
        private const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
        private const uint MOUSEEVENTF_RIGHTUP = 0x0010;
        private const uint MOUSEEVENTF_MIDDLEDOWN = 0x0020;
        private const uint MOUSEEVENTF_MIDDLEUP = 0x0040;

        private static void PerformLeftClickAtPoint(System.Drawing.Point point)
        {
            if (ClickOverride != null)
            {
                try { ClickOverride(point); } catch { }
                return;
            }
            System.Drawing.Point previous;
            try { GetCursorPos(out previous); } catch { previous = System.Windows.Forms.Cursor.Position; }

            try
            {
                var targetWindow = WindowFromPoint(point);
                if (targetWindow != IntPtr.Zero)
                {
                    SetForegroundWindow(targetWindow);
                }
            }
            catch { }

            SetCursorPos(point.X, point.Y);
            System.Threading.Thread.Sleep(35);

            var usedInputSimulator = false;
            try
            {
                var sim = new WindowsInput.InputSimulator();
                sim.Mouse.LeftButtonDown();
                System.Threading.Thread.Sleep(25);
                sim.Mouse.LeftButtonUp();
                usedInputSimulator = true;
            }
            catch { }

            if (!usedInputSimulator)
            {
                mouse_event(MOUSEEVENTF_LEFTDOWN, point.X, point.Y, 0, 0);
                System.Threading.Thread.Sleep(25);
                mouse_event(MOUSEEVENTF_LEFTUP, point.X, point.Y, 0, 0);
            }

            System.Threading.Thread.Sleep(70);
            try { SetCursorPos(previous.X, previous.Y); } catch { }
        }


        // InterpretAsync implementation
        public System.Threading.Tasks.Task<ActionBase?> InterpretAsync(string text)

        {
            text = (text ?? string.Empty).ToLowerInvariant().Trim();
            // Remove polite modifiers and extra punctuation
            text = RemovePoliteModifiers(text);
            text = WordReplacementLoader.Apply(text);
            text = text.Replace("  ", " ").Replace(".", "").Replace(",", "").Trim();

            // Common synonym / misrecognition: accept "quick clips" and "quick links" as "quick clicks"
            if (text.IndexOf("quick clips", StringComparison.InvariantCultureIgnoreCase) >= 0)
                text = System.Text.RegularExpressions.Regex.Replace(text, @"quick\s+clips", "quick clicks", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            if (text.IndexOf("quick links", StringComparison.InvariantCultureIgnoreCase) >= 0)
                text = System.Text.RegularExpressions.Regex.Replace(text, @"quick\s+links", "quick clicks", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            // Remove extra words that often appear in these commands
            var extraWords = new[] { "of this", "of others", "of other windows", "on top of others", "on top of this" };
            foreach (var ew in extraWords) text = text.Replace(ew, "");
            text = text.Trim();
            AppendLog($"[DEBUG] InterpretAsync normalized input: {text}\n");

            // ---- special debug: cloud vision request ----
            if (text.StartsWith("cloud vision ", StringComparison.OrdinalIgnoreCase))
            {
                var phrase = text.Substring("cloud vision ".Length).Trim();
                AppendLog($"[DEBUG] InterpretAsync matched CloudVision debug action: {phrase}\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(new VisualCloudAction(phrase));
            }
            if (text.StartsWith("vision ai ", StringComparison.OrdinalIgnoreCase))
            {
                var phrase = text.Substring("vision ai ".Length).Trim();
                AppendLog($"[DEBUG] InterpretAsync matched VisionAI debug action: {phrase}\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(new VisualCloudAction(phrase));
            }

            // Focus window by name: /focus [window name], focus [window name], focus window [name]
            string? focusPrefix = null;
            if (text.StartsWith("/focus ")) focusPrefix = "/focus ";
            else if (text.StartsWith("focus window ")) focusPrefix = "focus window ";
            else if (text.StartsWith("focus ")) focusPrefix = "focus ";
            if (focusPrefix != null)
            {
                var windowName = text.Substring(focusPrefix.Length).Trim();
                if (!string.IsNullOrEmpty(windowName))
                {
                    AppendLog($"[DEBUG] InterpretAsync matched FocusWindow command: {windowName}\n");
                    return System.Threading.Tasks.Task.FromResult<ActionBase?>(new NaturalCommands.FocusWindowAction(windowName));
                }
            }

            // Check popular commands override
            if (PopularCommands.TryGetValue(text, out var popularAction))
            {
                AppendLog($"[DEBUG] InterpretAsync matched PopularCommand: {text} -> {popularAction.GetType().Name}\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(popularAction);
            }

            // Handle refresh Visual Studio shortcuts command
            if (text.Contains("refresh visual studio shortcuts") || text.Contains("reload visual studio shortcuts") || text.Contains("update visual studio shortcuts"))
            {
                NaturalCommands.Helpers.VisualStudioShortcutHelper.RefreshShortcuts();
                AppendLog("[INFO] Refreshed Visual Studio shortcuts from .vssettings file\n");
                NaturalCommands.TrayNotificationHelper.ShowNotification("Shortcuts Refreshed", "Visual Studio keyboard shortcuts have been reloaded.", 5000);
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(null);
            }

            // Handle refresh Talon command catalog
            if (text.Contains("refresh talon commands") || text.Contains("reload talon commands") || text.Contains("update talon commands"))
            {
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(new RefreshTalonCatalogAction());
            }

            // Explicit help/command list queries
            var helpQueriesExact = new[] {
                    "what can i say", "help", "show commands", "show available commands", "list commands", "show help", "commands list", "available commands"
                };
            if (helpQueriesExact.Any(q => text.Equals(q, StringComparison.InvariantCultureIgnoreCase)))
            {
                // ShowAvailableCommands performs the appropriate UI (dialog or notification).
                // We return null here so no further ShowHelpAction is executed (avoids duplicate notifications).
                ShowAvailableCommands();
                AppendLog($"[DEBUG] InterpretAsync matched: ShowAvailableCommands displayed (help query)\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(null);
            }

            // Check for configured multi-action commands (exact match or normalized match)
            try
            {
                if (NaturalCommands.Helpers.MultiActionLoader.Commands.TryGetValue(text, out var multi) ||
                    NaturalCommands.Helpers.MultiActionLoader.Commands.TryGetValue(NaturalCommands.Helpers.MultiActionLoader.NormalizeKey(text), out multi))
                {
                    AppendLog($"[DEBUG] InterpretAsync matched multi-action command: {multi.Name}\n");
                    return System.Threading.Tasks.Task.FromResult<ActionBase?>(multi);
                }
            }
            catch { }

            // More robust matching for 'always on top'/'float above'/'restore' commands
            var alwaysOnTopPatterns = new[] {
                    "always on top", "on top", "float above", "float this window", "float above other windows",
                    "make this window float above", "make this window float", "float this window above",
                    "float window above", "make window float", "make window always on top",
                    "put this window on top", "put window on top", "make window float above", "put window above",
                    "float this window above other windows", "float window above other windows", "float window above others",
                    "float this window above others", "float window above",
                    "put this window above other windows", "put this window above others", "put window above other windows",
                    "put window above others", "make this window always on top", "make window always on top"
                };
            // Restore window (un-maximize)
            if ((text.Contains("restore") || text.Contains("unmaximize")) && text.Contains("window"))
            {
                var action = new MoveWindowAction(
                    Target: "active",
                    Monitor: "current",
                    Position: "center",
                    WidthPercent: 80,
                    HeightPercent: 80
                );
                AppendLog("Window maximized\n");
                AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name} (restore window)\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }
            bool matchedAlwaysOnTop = false;
            foreach (var pattern in alwaysOnTopPatterns)
            {
                if (text.Contains(pattern))
                {
                    matchedAlwaysOnTop = true;
                    AppendLog($"[DEBUG] InterpretAsync matched pattern: {pattern}\n");
                    break;
                }
            }
            // Also match regex variants like 'float.*window.*top' or 'make.*window.*top'
            if (!matchedAlwaysOnTop)
            {
                var regexPatterns = new[] {
                        "float.*window.*top", "make.*window.*top", "float.*window.*above", "make.*window.*float", "put.*window.*top", "put.*window.*above"
                    };
                foreach (var rx in regexPatterns)
                {
                    if (System.Text.RegularExpressions.Regex.IsMatch(text, rx))
                    {
                        matchedAlwaysOnTop = true;
                        AppendLog($"[DEBUG] InterpretAsync matched regex: {rx}\n");
                        break;
                    }
                }
                // Catch-all: match any phrase containing 'float', 'window', and 'above' in any order
                if (!matchedAlwaysOnTop)
                {
                    var words = new[] { "float", "window", "above" };
                    bool allPresent = words.All(w => text.Contains(w));
                    if (allPresent)
                    {
                        matchedAlwaysOnTop = true;
                        AppendLog("[DEBUG] InterpretAsync matched catch-all: float/window/above\n");
                    }
                }
            }
            if (matchedAlwaysOnTop)
            {
                // The always-on-top feature is disabled because it was causing accidental
                // behaviors for users. Log and return null so no action is executed.
                AppendLog("[INFO] InterpretAsync: 'always on top' command detected but feature is disabled by configuration.\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(null);
            }
            // Send key sequences
            if (text.StartsWith("press "))
            {
                var keysText = text.Substring(6).Trim();
                var action = new SendKeysAction(keysText);
                AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name} (send keys)\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }
            // Maximize/full screen window
            if ((text.Contains("maximize") || text.Contains("full screen")) && text.Contains("window"))
            {
                var action = new MoveWindowAction(
                    Target: "active",
                    Monitor: "current",
                    Position: "center",
                    WidthPercent: 100,
                    HeightPercent: 100
                );
                NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync matched: {action.GetType().Name}");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }
            // Move window to other monitor (next)
            if ((text.Contains("move") || text.Contains("snap")) && text.Contains("window") && (text.Contains("other monitor") || text.Contains("next monitor") || text.Contains("other screen") || text.Contains("next screen") || text.Contains("my other monitor")))
            {
                var action = new MoveWindowAction(
                    Target: "active",
                    Monitor: "next",
                    Position: null,
                    WidthPercent: 0,
                    HeightPercent: 0
                );
                NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync matched: {action.GetType().Name} (next monitor)");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }
            // Move window to left half (robust)
            if ((text.Contains("left half") || (text.Contains("left") && text.Contains("half"))) || ((text.Contains("move") || text.Contains("snap")) && text.Contains("window") && text.Contains("left")))
            {
                var action = new MoveWindowAction(
                    Target: "active",
                    Monitor: "current",
                    Position: "left",
                    WidthPercent: 50,
                    HeightPercent: 100
                );
                NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync matched: {action.GetType().Name} (left half)");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }
            // Move window to right half (robust)
            if ((text.Contains("right half") || (text.Contains("right") && text.Contains("half"))) || ((text.Contains("move") || text.Contains("snap")) && text.Contains("window") && text.Contains("right")))
            {
                var action = new MoveWindowAction(
                    Target: "active",
                    Monitor: "current",
                    Position: "right",
                    WidthPercent: 50,
                    HeightPercent: 100
                );
                NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync matched: {action.GetType().Name} (right half)");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }
            // Open My Computer / This PC (robust)
            if (text.Contains("open my computer") || text.Contains("open this pc") || text.Contains("open my pc") || (text.Contains("open") && text.Contains("computer")) || (text.Contains("open") && text.Contains("pc")))
            {
                var action = new OpenFolderAction("MyComputer");
                NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync matched: {action.GetType().Name} (my computer)");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }
            // Open documents folder (robust)
            if (text.Contains("open documents") || (text.Contains("open") && text.Contains("document")))
            {
                var action = new OpenFolderAction("Documents");
                NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync matched: {action.GetType().Name} (documents)");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }
            // Open downloads folder (robust)
            if (text.Contains("open downloads") || (text.Contains("open") && text.Contains("download")))
            {
                var action = new OpenFolderAction("Downloads");
                NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync matched: {action.GetType().Name} (downloads)");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }

            // Open settings explicitly: "open settings", "settings", "show settings"
            if (text.Equals("open settings", StringComparison.InvariantCultureIgnoreCase)
                || text.Equals("settings", StringComparison.InvariantCultureIgnoreCase)
                || text.Equals("show settings", StringComparison.InvariantCultureIgnoreCase))
            {
                var action = new OpenSettingsAction();
                AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name} (settings)\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }

            // Windows Terminal specific commands
            string? wtProcName = NaturalCommands.CurrentApplicationHelper.GetCurrentProcessName();
            if (wtProcName == "windowsterminal" || wtProcName == "WindowsTerminal")
            {
                // First check for Windows Terminal keyboard shortcuts
                if (CommandDefinitions.WindowsTerminalShortcuts.TryGetValue(text, out var shortcut))
                {
                    var action = new WindowsTerminalShortcutAction(shortcut, text);
                    NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync matched Windows Terminal command: '{text}' -> '{shortcut}'");
                    return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
                }

                // Normalize text for dotnet commands - speech recognition often transcribes "dotnet" as "dot net" or "don't net"
                string normalizedDotnetText = text
                    .Replace("dot net", "dotnet")
                    .Replace("don't net", "dotnet")
                    .Replace("dont net", "dotnet")
                    .Replace("dot-net", "dotnet")
                    .Replace("dot_net", "dotnet");

                // Check for dotnet CLI commands (try both original and normalized text)
                if (CommandDefinitions.DotNetCommandMappings.TryGetValue(text, out var dotnetCmd) ||
                    CommandDefinitions.DotNetCommandMappings.TryGetValue(normalizedDotnetText, out dotnetCmd))
                {
                    var action = new RunTerminalCommandAction(dotnetCmd.TerminalCommand, dotnetCmd.Description);
                    NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync matched dotnet command: '{text}' (normalized: '{normalizedDotnetText}') -> '{dotnetCmd.TerminalCommand}'");
                    return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
                }
            }

            // Windows Explorer specific commands
            if (wtProcName == "explorer")
            {
                // Check for Windows Explorer shortcuts (includes UI Automation commands with "uia:" prefix)
                if (CommandDefinitions.WindowsExplorerShortcuts.TryGetValue(text, out var explorerShortcut))
                {
                    var action = new WindowsExplorerShortcutAction(explorerShortcut, text);
                    NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync matched Windows Explorer command: '{text}' -> '{explorerShortcut}'");
                    return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
                }
            }

            // Open mapped applications (expanded) 
            // Use WebsiteNavigator for website navigation commands
            if (WebsiteNavigator.TryParseWebsiteCommand(text, out var url))
            {
                if (!string.IsNullOrWhiteSpace(url))
                {
                    var action = new OpenWebsiteAction(url);
                    AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name} (website: {url})\n");
                    return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
                }
            }
            // Open mapped applications (expanded)
            if (text.StartsWith("open "))
            {
                var appName = text.Substring(5).Trim();
                // Normalize app name (remove filler words as whole words only)
                appName = Regex.Replace(appName, "\\bthe\\b", "", RegexOptions.IgnoreCase).Trim();
                appName = Regex.Replace(appName, "\\bapplication\\b", "", RegexOptions.IgnoreCase).Trim();
                appName = Regex.Replace(appName, "\\bapp\\b", "", RegexOptions.IgnoreCase).Trim();
                appName = Regex.Replace(appName, "\\s+", " ").Trim();
                var fillerTokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "up", "for", "me", "now", "please", "just"
                };
                var filteredTokens = appName
                    .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                    .Where(t => !fillerTokens.Contains(t))
                    .ToArray();
                if (filteredTokens.Length > 0)
                {
                    appName = string.Join(" ", filteredTokens);
                }
                // Strip vendor prefixes like "microsoft " or "ms " to handle phrases such as "microsoft paint"
                if (appName.StartsWith("microsoft ", StringComparison.OrdinalIgnoreCase))
                    appName = appName.Substring("microsoft ".Length).Trim();
                else if (appName.StartsWith("ms ", StringComparison.OrdinalIgnoreCase))
                    appName = appName.Substring("ms ".Length).Trim();
                // Special case: "terminal" or "windows terminal" or "cmd" or "command prompt"
                if (appName == "terminal" || appName == "windows terminal" || appName == "cmd" || appName == "command prompt")
                    appName = "terminal";
                if (AppMappings.TryGetValue(appName, out var exe))
                {
                    var action = new LaunchAppAction(exe);
                    AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name} (mapped app: {appName} -> {exe})\n");
                    return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
                }
                var embeddedAppName = AppMappings.Keys
                    .OrderByDescending(k => k.Length)
                    .FirstOrDefault(k => Regex.IsMatch(appName, $"\\b{Regex.Escape(k)}\\b", RegexOptions.IgnoreCase));
                if (!string.IsNullOrWhiteSpace(embeddedAppName) && AppMappings.TryGetValue(embeddedAppName, out exe))
                {
                    var action = new LaunchAppAction(exe);
                    AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name} (embedded mapped app: {embeddedAppName} -> {exe})\n");
                    return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
                }
                // Fallback: show Alt+Tab switcher and hold Alt
                var fallbackAction = new LaunchAppAction("focus-fallback");
                AppendLog($"[DEBUG] InterpretAsync fallback: {fallbackAction.GetType().Name} (focus-fallback for: {appName})\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(fallbackAction);
            }
            // "type ..." maps to SendKeysAction
            if (text.StartsWith("type "))
            {
                var keysText = text.Substring(5).Trim();
                var action = new SendKeysAction(keysText);
                AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name} (type keys)\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }
            // Emoji commands
            // "emoji set <name> <emoji>" -> set mapping
            if (text.StartsWith("emoji set "))
            {
                // Examples: "emoji set happy 😀" or "emoji set happy :D"
                var rest = text.Substring("emoji set ".Length).Trim();
                if (!string.IsNullOrEmpty(rest))
                {
                    var parts = rest.Split(new[] { ' ' }, 2, StringSplitOptions.RemoveEmptyEntries);
                    if (parts.Length == 2)
                    {
                        var name = parts[0].Trim();
                        var emoji = parts[1].Trim();
                        EmojiManager.SetCommandEmoji(name, emoji);
                        AppendLog($"[DEBUG] InterpretAsync: Set emoji mapping: {name} -> {emoji}\n");
                        return System.Threading.Tasks.Task.FromResult<ActionBase?>(new EmojiAction(name, emoji));
                    }
                }
            }
            // "emoji type <emoji>" -> type this emoji immediately
            if (text.StartsWith("emoji type "))
            {
                var emojiText = text.Substring("emoji type ".Length).Trim();
                if (!string.IsNullOrEmpty(emojiText))
                {
                    var action = new EmojiAction(null, emojiText);
                    AppendLog($"[DEBUG] InterpretAsync: Emoji type action for: {emojiText}\n");
                    return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
                }
            }
            // "emoji <name>" -> type configured emoji for name
            if (text.StartsWith("emoji ") && !text.StartsWith("emoji set") && !text.StartsWith("emoji type"))
            {
                var name = text.Substring("emoji ".Length).Trim();
                if (!string.IsNullOrEmpty(name))
                {
                    var emoji = EmojiManager.GetCommandEmoji(name);
                    if (!string.IsNullOrEmpty(emoji))
                    {
                        var action = new EmojiAction(name, emoji);
                        AppendLog($"[DEBUG] InterpretAsync: Emoji name action for: {name} -> {emoji}\n");
                        return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
                    }
                    else
                    {
                        // No mapping found; respond with AI fallback or suggest how to set
                        AppendLog($"[DEBUG] InterpretAsync: No emoji mapping for: {name}\n");
                        // Suggest set command via ShowAvailableCommands or let AI handle
                        return System.Threading.Tasks.Task.FromResult<ActionBase?>(null);
                    }
                }
            }
            // Supported apps for close tab
            if (text.Trim().Equals("close tab", StringComparison.InvariantCultureIgnoreCase) || text.Trim().Equals("closed tab", StringComparison.InvariantCultureIgnoreCase))
            {
                string? procName = NaturalCommands.CurrentApplicationHelper.GetCurrentProcessName();
                if (!string.IsNullOrEmpty(procName) && SupportedCloseTabApps.Contains(procName))
                {
                    var closeTabAction = new CloseTabAction();
                    AppendLog($"[DEBUG] InterpretAsync: Rule-based match for 'close tab' in supported app: {procName}\n");
                    return System.Threading.Tasks.Task.FromResult<ActionBase?>(closeTabAction);
                }
            }
            // Visual Studio code search (Go to All) shortcut
            var codeSearchPatterns = new[] {
                "search code", "code search", "find code", "open code search", "search for code"
            };
            if (codeSearchPatterns.Any(p => text.Contains(p)))
            {
                var action = new SendKeysAction("control ,");
                AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name} (code search)\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }

            // Show letters navigation feature
            var showLettersPatterns = new[] {
                "show letters", "natural show letters", "display letters", "label elements", 
                "show labels", "click by letter", "letter navigation"
            };
            if (showLettersPatterns.Any(p => text.Contains(p)))
            {
                var action = new ShowLettersAction(ScopeToActiveWindow: true);
                AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name} (show letters)\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }

            var showTaskbarPatterns = new[] {
                "show taskbar", "taskbar letters", "show tray", "focus taskbar"
            };
            if (showTaskbarPatterns.Any(p => text.Contains(p)))
            {
                var action = new ShowTaskbarAction();
                AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name} (show taskbar)\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }

            var showDesktopPatterns = new[] {
                "show desktop", "show the desktop", "desktop letters", "show desktop letters", "show desktop icons", "desktop icons", "focus desktop"
            };
            if (showDesktopPatterns.Any(p => text.Contains(p)))
            {
                var action = new ShowDesktopAction();
                AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name} (show desktop)\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }

            if (text.StartsWith("identify ", StringComparison.InvariantCultureIgnoreCase)
                || text.StartsWith("id ", StringComparison.InvariantCultureIgnoreCase))
            {
                var prefix = text.StartsWith("identify ", StringComparison.InvariantCultureIgnoreCase) ? "identify " : "id ";
                var targetPhrase = text.Substring(prefix.Length).Trim();
                if (!string.IsNullOrWhiteSpace(targetPhrase))
                {
                    var action = new VisualIdentifyClickAction(targetPhrase);
                    AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name} (target={targetPhrase})\n");
                    return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
                }
            }

            if (text.Equals("show candidates", StringComparison.InvariantCultureIgnoreCase)
                || text.Equals("show visual candidates", StringComparison.InvariantCultureIgnoreCase))
            {
                var action = new VisualShowCandidatesAction();
                AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name}\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }

            var chooseMatch = Regex.Match(text, @"^(choose|select|pick)\s+(?:number\s+)?([a-z0-9]+(?:st|nd|rd|th)?)\s*$", RegexOptions.IgnoreCase);
            if (chooseMatch.Success && Helpers.SpokenNumberParser.TryParsePositive(chooseMatch.Groups[2].Value, out var chosenNumber))
            {
                var action = new VisualChooseCandidateAction(chosenNumber);
                AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name} (candidate={chosenNumber})\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }

            // Quick Clicks control phrases removed (feature deprecated).

            // Quick Clicks dynamic region parsing removed (feature deprecated).

            // Mouse movement commands
            // Support both "move left" and "mouse move left" or "mouse left"
            string? direction = null;
            if (text.StartsWith("move "))
            {
                direction = text.Substring(5).Trim();
            }
            else if (text.StartsWith("mouse move "))
            {
                direction = text.Substring(11).Trim();
            }
            else if (text.StartsWith("mouse ") && !text.StartsWith("mouse stop") && !text.StartsWith("mouse faster") && !text.StartsWith("mouse slower"))
            {
                // "mouse left" -> "left"
                direction = text.Substring(6).Trim();
            }
            
            if (direction != null)
            {
                var action = new StartMouseMoveAction(direction);
                AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name} (move mouse {direction})\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }
            
            if (text == "stop mouse" || text == "mouse stop")
            {
                var action = new StopMouseMoveAction(PerformClick: false);
                AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name} (stop mouse)\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }
            if (text == "stop click" || text == "mouse stop click")
            {
                var action = new StopMouseMoveAction(PerformClick: true, IsRightClick: false);
                AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name} (stop and click)\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }
            if (text == "stop right click" || text == "stop and right click" || text == "mouse stop right click")
            {
                var action = new StopMouseMoveAction(PerformClick: true, IsRightClick: true);
                AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name} (stop and right click)\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }
            if (text == "faster" || text == "slower" || text == "mouse faster" || text == "mouse slower")
            {
                // Extract the actual speed adjustment (faster/slower)
                string speedAdjustment = text.Replace("mouse ", "").Trim();
                var action = new AdjustMouseSpeedAction(speedAdjustment);
                AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name} (adjust speed {speedAdjustment})\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }

            // Auto-click commands
            if (text == "auto click" || text == "enable auto click" || text == "start auto click")
            {
                var action = new StartAutoClickAction(DelayMs: 0); // 0 = use default from settings
                AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name} (enable auto-click)\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }
            if (text == "stop auto click" || text == "disable auto click" || text == "auto click off")
            {
                var action = new StopAutoClickAction();
                AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name} (disable auto-click)\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }

            // Adjust auto-click delay via voice: "auto click faster" / "auto click slower"
            if (text == "auto click faster" || text == "auto click speed up" || text == "auto click speedup")
            {
                var action = new AdjustAutoClickDelayAction("faster");
                AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name} (auto-click faster)\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }
            if (text == "auto click slower" || text == "auto click slow down" || text == "auto click slowdown")
            {
                var action = new AdjustAutoClickDelayAction("slower");
                AppendLog($"[DEBUG] InterpretAsync matched: {action.GetType().Name} (auto-click slower)\n");
                return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
            }

            // Visual Studio Command Lookup

            if (IsVisualStudioActive())
            {
                // First, check explicit tool window mappings
                foreach (var kvp in VisualStudioToolWindowMappings)
                {
                    if (text.Equals(kvp.Key, StringComparison.InvariantCultureIgnoreCase) || text.Contains(kvp.Key, StringComparison.InvariantCultureIgnoreCase))
                    {
                        var mappedAction = new ExecuteVSCommandAction(kvp.Value);
                        AppendLog($"[DEBUG] InterpretAsync matched tool window mapping: {kvp.Key} -> {kvp.Value}\n");
                        return System.Threading.Tasks.Task.FromResult<ActionBase?>(mappedAction);
                    }
                }
                // Then, check explicit canonical mappings for common VS commands. This avoids
                // choosing context-menu or other exported commands that are not the canonical ones.
                foreach (var kvp in VisualStudioCanonicalMappings)
                {
                    if (text.Equals(kvp.Key, StringComparison.InvariantCultureIgnoreCase) || text.Contains(kvp.Key, StringComparison.InvariantCultureIgnoreCase))
                    {
                        var mappedAction = new ExecuteVSCommandAction(kvp.Value);
                        AppendLog($"[DEBUG] InterpretAsync matched canonical VS mapping: {kvp.Key} -> {kvp.Value}\n");
                        return System.Threading.Tasks.Task.FromResult<ActionBase?>(mappedAction);
                    }
                }

                // Ensure commands are loaded
                if (NaturalCommands.Helpers.VisualStudioCommandLoader.GetCommands().Count == 0)
                {
                    string jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "vs_commands.json");
                    if (!File.Exists(jsonPath))
                    {
                         // Try looking up three levels (project root during dev)
                         jsonPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "vs_commands.json");
                    }
                    NaturalCommands.Helpers.VisualStudioCommandLoader.LoadCommands(jsonPath);
                }

                var vsCommand = NaturalCommands.Helpers.VisualStudioCommandLoader.FindCommand(text);
                if (vsCommand != null)
                {
                     var action = new ExecuteVSCommandAction(vsCommand.Name);
                     NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync matched VS Command: {vsCommand.Name}");
                     return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
                }
            }

            // Fallback for unhandled commands: log and let HandleNaturalAsync perform AI/Talon routing
            NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync: No rule-based match for: {text}");
            return System.Threading.Tasks.Task.FromResult<ActionBase?>(null);
            // End of InterpretAsync
        }
        private static readonly string[] SupportedCloseTabApps = new[] { "chrome", "msedge", "firefox", "brave", "opera", "code", "devenv" };

                    public string ExecuteActionAsync(ActionBase action)
                    {
                        AppendLog($"[DEBUG] ExecuteActionAsync: Action type: {(action == null ? "null" : action.GetType().Name)}\n");
                        AppendLog($"[DEBUG] ExecuteActionAsync: action.GetType().FullName: {(action == null ? "null" : action.GetType().FullName)}\n");
                        AppendLog($"[DEBUG] ExecuteActionAsync: Checking if action is MoveWindowAction\n");
                        AppendLog($"[DEBUG] ExecuteActionAsync: action.GetType().AssemblyQualifiedName: {(action == null ? "null" : action.GetType().AssemblyQualifiedName)}\n");
                        NaturalCommands.Helpers.Logger.EnsureLogDirExists();
            if (action is MoveWindowAction move)
            {
                // Delegate to WindowManager
                return NaturalCommands.Helpers.WindowManager.ExecuteMoveWindow(move);
            }
            // All window movement logic is now inside the MoveWindowAction block above
            else if (action is CloseTabAction)
            {
                // Always log 'Sent Ctrl+W' for close tab attempts, even if app is unsupported or process name is missing
                AppendLog("Sent Ctrl+W\n");
                string? procName = CurrentApplicationHelper.GetCurrentProcessName();
                if (string.IsNullOrEmpty(procName))
                {
                    return "Could not detect current application.";
                }

                if (SupportedCloseTabApps.Contains(procName))
                {
                    try
                    {
                        var sim = new WindowsInput.InputSimulator();
                        if (procName == "devenv")
                        {
                            // In Visual Studio, use DTE to close the document window
                            try
                            {
                                bool success = NaturalCommands.Helpers.VisualStudioHelper.ExecuteCommand("Window.CloseDocumentWindow");
                                if (success)
                                {
                                    NaturalCommands.Helpers.Logger.LogDebug("Executed Window.CloseDocumentWindow via DTE in Visual Studio.");
                                    return "Executed Window.CloseDocumentWindow via DTE.";
                                }
                                else
                                {
                                    // Fallback to Ctrl+F4 if DTE fails
                                    sim.Keyboard.ModifiedKeyStroke(WindowsInput.Native.VirtualKeyCode.CONTROL, WindowsInput.Native.VirtualKeyCode.F4);
                                    NaturalCommands.Helpers.Logger.LogError("DTE failed, sent Ctrl+F4 to Visual Studio.");
                                    return "DTE failed, sent Ctrl+F4.";
                                }
                            }
                            catch (Exception ex)
                            {
                                NaturalCommands.Helpers.Logger.LogError($"Failed to execute VS command: {ex.Message}");
                                return $"Failed to execute VS command: {ex.Message}";
                            }
                        }
                        sim.Keyboard.ModifiedKeyStroke(WindowsInput.Native.VirtualKeyCode.CONTROL, WindowsInput.Native.VirtualKeyCode.VK_W);
                        return $"Sent Ctrl+W to {procName} (close tab).";
                    }
                    catch (Exception ex)
                    {
                        NaturalCommands.Helpers.Logger.LogError($"Failed to send Ctrl+W to {procName}: {ex.Message}");
                        return $"Failed to send Ctrl+W to {procName}: {ex.Message}";
                    }
                }
                else
                {
                    return $"Current app '{procName}' is not supported for 'close tab'.";
                }
            }
            // Open voice dictation form action: show form, process commands as they're sent
            else if (action is OpenVoiceDictationFormAction dictAction)
            {
                try
                {
                    string lastResult = "Voice dictation form opened.";
                    NaturalCommands.Helpers.VoiceDictationHelper.ShowVoiceDictation(dictAction.TimeoutMs, (text) =>
                    {
                        try
                        {
                            if (string.IsNullOrWhiteSpace(text))
                                return;
                            // Apply word replacements
                            text = NaturalCommands.Helpers.WordReplacementHelper.ApplyWordReplacements(text);
                            // Interpret the dictated text and execute resulting action(s)
                            var interpretedTask = InterpretAsync(text);
                            var interpreted = interpretedTask?.Result;
                            if (interpreted != null)
                            {
                                lastResult = ExecuteActionAsync(interpreted);
                            }
                        }
                        catch (Exception ex)
                        {
                            AppendLog($"[ERROR] Command execution failed: {ex.Message}\n");
                        }
                    });
                    return lastResult;
                }
                catch (Exception ex)
                {
                    return $"Voice dictation failed: {ex.Message}";
                }
            }
            else if (action is OpenSettingsAction)
            {
                try
                {
                    // Show modal settings dialog. If this fails, return a descriptive error.
                    var settingsForm = new SettingsForm();
                    settingsForm.ShowDialog();
                    return "Settings displayed.";
                }
                catch (Exception ex)
                {
                    AppendLog($"[ERROR] Failed to open settings form: {ex.Message}\n");
                    return $"Failed to open settings: {ex.Message}";
                }
            }
            // Mouse movement actions
            else if (action is StartMouseMoveAction moveAction)
            {
                return NaturalCommands.Helpers.MouseMoveManager.StartMoving(moveAction.Direction);
            }
            else if (action is StopMouseMoveAction stopAction)
            {
                return NaturalCommands.Helpers.MouseMoveManager.StopMoving(stopAction.PerformClick, stopAction.IsRightClick);
            }
            else if (action is AdjustMouseSpeedAction speedAction)
            {
                return NaturalCommands.Helpers.MouseMoveManager.AdjustSpeed(speedAction.SpeedChange);
            }
            // Auto-click actions
            else if (action is StartAutoClickAction autoClickAction)
            {
                AppendLog($"[DEBUG] ExecuteActionAsync: Executing StartAutoClickAction\n");
                return NaturalCommands.Helpers.AutoClickManager.Start(autoClickAction.DelayMs);
            }
            else if (action is StopAutoClickAction)
            {
                AppendLog($"[DEBUG] ExecuteActionAsync: Executing StopAutoClickAction\n");
                NaturalCommands.Helpers.Logger.LogDebug("ExecuteActionAsync: About to call AutoClickManager.Stop()");
                var result = NaturalCommands.Helpers.AutoClickManager.Stop();
                NaturalCommands.Helpers.Logger.LogDebug($"ExecuteActionAsync: AutoClickManager.Stop() returned: {result}");
                return result;
            }
            // Windows Terminal shortcut action
            else if (action is WindowsTerminalShortcutAction wtAction)
            {
                AppendLog($"[DEBUG] ExecuteActionAsync: Executing WindowsTerminalShortcutAction ('{wtAction.CommandText}' -> '{wtAction.Shortcut}')\n");
                NaturalCommands.Helpers.Logger.LogDebug($"ExecuteActionAsync: Windows Terminal shortcut: '{wtAction.CommandText}' -> '{wtAction.Shortcut}'");
                var result = NaturalCommands.Helpers.KeySender.SendShortcut(wtAction.Shortcut);
                return $"Windows Terminal: {wtAction.CommandText} ({result})";
            }
            // Windows Explorer shortcut action
            else if (action is WindowsExplorerShortcutAction explorerAction)
            {
                AppendLog($"[DEBUG] ExecuteActionAsync: Executing WindowsExplorerShortcutAction ('{explorerAction.CommandText}' -> '{explorerAction.Shortcut}')\n");
                NaturalCommands.Helpers.Logger.LogDebug($"ExecuteActionAsync: Windows Explorer shortcut: '{explorerAction.CommandText}' -> '{explorerAction.Shortcut}'");
                
                // Small delay to ensure Explorer window has focus before sending shortcut
                System.Threading.Thread.Sleep(100);
                
                // Check if this is a UI Automation command (prefix "uia:")
                if (explorerAction.Shortcut.StartsWith("uia:", StringComparison.OrdinalIgnoreCase))
                {
                    var uiaCommand = explorerAction.Shortcut.Substring(4).ToLowerInvariant();
                    bool success = false;
                    
                    switch (uiaCommand)
                    {
                        case "view":
                            success = NaturalCommands.Helpers.UIAutomationHelper.ClickExplorerViewButton();
                            break;
                        case "sort":
                            success = NaturalCommands.Helpers.UIAutomationHelper.ClickExplorerSortButton();
                            break;
                        case "new":
                            success = NaturalCommands.Helpers.UIAutomationHelper.ClickExplorerNewButton();
                            break;
                        default:
                            NaturalCommands.Helpers.Logger.LogError($"Unknown UI Automation command: {uiaCommand}");
                            break;
                    }
                    
                    if (success)
                    {
                        return $"Windows Explorer: Opened {uiaCommand} menu";
                    }
                    else
                    {
                        return $"Windows Explorer: Failed to open {uiaCommand} menu - element not found";
                    }
                }
                // Check if shortcut contains commas (menu sequence like "shift+f10,v,x")
                else if (explorerAction.Shortcut.Contains(','))
                {
                    // This is a menu sequence - send each key separately with delays
                    var keys = explorerAction.Shortcut.Split(',').Select(k => k.Trim()).ToArray();
                    foreach (var key in keys)
                    {
                        NaturalCommands.Helpers.KeySender.SendShortcut(key);
                        System.Threading.Thread.Sleep(200); // Wait for menu to open
                    }
                    return $"Windows Explorer: {explorerAction.CommandText} (menu sequence)";
                }
                else
                {
                    var result = NaturalCommands.Helpers.KeySender.SendShortcut(explorerAction.Shortcut);
                    return $"Windows Explorer: {explorerAction.CommandText} ({result})";
                }
            }
            // Run terminal command action (types command and presses Enter)
            else if (action is RunTerminalCommandAction terminalCmd)
            {
                AppendLog($"[DEBUG] ExecuteActionAsync: Executing RunTerminalCommandAction ('{terminalCmd.Command}')\n");
                NaturalCommands.Helpers.Logger.LogDebug($"ExecuteActionAsync: Running terminal command: '{terminalCmd.Command}' - {terminalCmd.Description}");
                try
                {
                    var sim = new WindowsInput.InputSimulator();
                    // Type the command
                    sim.Keyboard.TextEntry(terminalCmd.Command);
                    // Small delay before pressing Enter
                    System.Threading.Thread.Sleep(50);
                    // Press Enter to execute
                    sim.Keyboard.KeyPress(WindowsInput.Native.VirtualKeyCode.RETURN);
                    NaturalCommands.Helpers.Logger.LogDebug($"ExecuteActionAsync: Typed and executed terminal command: '{terminalCmd.Command}'");
                    return $"Terminal: {terminalCmd.Description} ({terminalCmd.Command})";
                }
                catch (Exception ex)
                {
                    NaturalCommands.Helpers.Logger.LogError($"Failed to execute terminal command: {ex.Message}");
                    return $"Failed to execute terminal command: {ex.Message}";
                }
            }
            else if (action is RefreshTalonCatalogAction)
            {
                try
                {
                    var count = NaturalCommands.Helpers.TalonCommandCatalog.Reload();
                    return $"Refreshed Talon commands ({count} loaded).";
                }
                catch (Exception ex)
                {
                    return $"Failed to refresh Talon commands: {ex.Message}";
                }
            }
            else if (action is RunTalonCommandAction talonAction)
            {
                return NaturalCommands.Helpers.TalonCommandExecutor.Execute(talonAction);
            }
            else if (action is VisualCloudAction cloudAction)
            {
                // debug-only action: send screenshot+phrase to AI model without clicking
                var candidates = Helpers.VisualTargetingService.GetCloudVisionCandidates(cloudAction.TargetPhrase);
                Helpers.Logger.LogInfo($"ExecuteActionAsync VisualCloudAction: cloud returned {candidates.Count} candidates for '{cloudAction.TargetPhrase}'.");
                if (candidates.Count > 0)
                {
                    foreach (var c in candidates)
                    {
                        Helpers.Logger.LogInfo($"  candidate '{c.Label}' ({c.Confidence:0.00}) source={c.Source} reason={c.Reason}");
                    }
                }
                else
                {
                    Helpers.Logger.LogInfo("  cloud vision produced no candidates");
                }
                return $"Cloud returned {candidates.Count} candidates for '{cloudAction.TargetPhrase}'. See log for details.";
            }

            else if (action is VisualIdentifyClickAction visualIdentify)
            {
                try
                {
                    if (!AppSettings.Instance.VisualTargeting.Enabled)
                    {
                        return "Visual targeting is disabled in settings.";
                    }

                    System.Collections.Generic.List<VisualTargetCandidate> candidates;
                    if (VisualIdentifyCandidatesOverride != null)
                    {
                        candidates = VisualIdentifyCandidatesOverride(visualIdentify.TargetPhrase) ?? new System.Collections.Generic.List<VisualTargetCandidate>();
                    }
                    else
                    {
                        candidates = Helpers.VisualTargetingService.IdentifyCandidates(visualIdentify.TargetPhrase);
                    }
                    Helpers.Logger.LogInfo($"ExecuteActionAsync VisualIdentifyClickAction: got {candidates.Count} candidates for '{visualIdentify.TargetPhrase}'.");

                    if (candidates.Count == 0)
                    {
                        Helpers.VisualCandidateSessionStore.Clear();
                        VisualCandidateOverlayForm.HideOverlay();

                        // Show debug overlay with screenshot for troubleshooting
                        try
                        {
                            var fgBounds = WindowUtils.GetForegroundWindowBounds();
                            var screenshot = fgBounds != Rectangle.Empty
                                ? Helpers.ScreenCaptureService.CaptureRegionJpeg(fgBounds, maxLongEdge: 1400, quality: 60)
                                : Helpers.ScreenCaptureService.CaptureAllScreensJpeg(maxLongEdge: 1400, quality: 60);
                            var debugMsg = $"Searching for: '{visualIdentify.TargetPhrase}'\n" +
                                           $"No visual matches found.\n" +
                                           $"Try 'show letters' to see clickable UI elements.";
                            Image? img = null;
                            if (!string.IsNullOrEmpty(screenshot.ImageBase64Jpeg))
                            {
                                byte[] imageBytes = Convert.FromBase64String(screenshot.ImageBase64Jpeg);
                                using (var ms = new System.IO.MemoryStream(imageBytes))
                                {
                                    using (var tempImg = Image.FromStream(ms))
                                    {
                                        img = new Bitmap(tempImg);
                                    }
                                }
                            }
                            VisualTargetDebugOverlayForm.ShowDebug(visualIdentify.TargetPhrase, img, debugMsg, timeoutMs: 4000);
                        }
                        catch (Exception ex)
                        {
                            Helpers.Logger.LogError($"Failed to show debug overlay: {ex.Message}");
                        }

                        return $"No visual match found for '{visualIdentify.TargetPhrase}'.";
                    }

                    // Session and persisted ordering will be created only if we show the overlay
                    // (we want the overlay numbering to match the persisted session order).

                    bool shouldAutoClickTop = false;
                    string autoClickReason = string.Empty;
                    var confidenceThreshold = AppSettings.Instance.VisualTargeting.AutoClickConfidenceThreshold;
                    var workaroundMinConfidence = Math.Max(0.45, confidenceThreshold - 0.35);
                    Helpers.Logger.LogDebug($"Visual identify: {candidates.Count} candidates found, confidenceThreshold={confidenceThreshold:0.00}, workaroundMinConfidence={workaroundMinConfidence:0.00}.");
                    if (candidates.Count > 0)
                    {
                        Helpers.Logger.LogDebug($"Visual identify: top candidate is '{candidates[0].Label}' (confidence {candidates[0].Confidence:0.00}, source '{candidates[0].Source}').");
                    }

                    if (candidates.Count >= 1)
                    {
                        bool isSingle = candidates.Count == 1;
                        if (isSingle)
                        {
                            var top = candidates[0];
                            var topLabel = (top.Label ?? string.Empty).Trim();
                            var target = (visualIdentify.TargetPhrase ?? string.Empty).Trim();

                            // For OCR single candidates, require a tight textual match to avoid
                            // clicking long sentence fragments that merely contain the token.
                            var isOcr = string.Equals(top.Source, "ocr", StringComparison.OrdinalIgnoreCase);
                            var exactLabelMatch = topLabel.Equals(target, StringComparison.OrdinalIgnoreCase);
                            var startsWithWholeTarget = !string.IsNullOrWhiteSpace(target)
                                && topLabel.StartsWith(target + " ", StringComparison.OrdinalIgnoreCase);
                            var boundedLength = topLabel.Length <= Math.Max(24, target.Length + 12);

                            if (!isOcr || (top.Confidence >= 0.92 && (exactLabelMatch || (startsWithWholeTarget && boundedLength))))
                            {
                                shouldAutoClickTop = true;
                                autoClickReason = isOcr ? "single candidate (safe ocr)" : "single candidate";
                            }
                            else
                            {
                                Helpers.Logger.LogInfo($"Visual identify: single OCR candidate '{topLabel}' rejected for auto-click (target='{target}', confidence={top.Confidence:0.00}).");
                            }
                        }
                        else if (candidates[0].Confidence >= workaroundMinConfidence)
                        {
                            // Safer default for multiple candidates: require the top candidate to be from UIA before auto-clicking.
                            if (string.Equals(candidates[0].Source, "uia", StringComparison.OrdinalIgnoreCase))
                            {
                                shouldAutoClickTop = true;
                                var top = candidates[0];
                                var second = candidates.Count > 1 ? candidates[1] : null;
                                if (second == null)
                                {
                                    autoClickReason = "single candidate workaround (uia)";
                                }
                                else
                                {
                                    autoClickReason = $"top candidate workaround (uia) ({top.Confidence:0.00} vs {second.Confidence:0.00})";
                                }
                            }
                            else
                            {
                                Helpers.Logger.LogInfo($"Visual identify: top candidate source is '{candidates[0].Source}' - skipping auto-click to avoid OCR-only auto-clicks for multi-candidate results.");
                            }
                        }
                        else
                        {
                            Helpers.Logger.LogInfo($"Visual identify: top candidate confidence {candidates[0].Confidence:0.00} below required {workaroundMinConfidence:0.00}; not auto-clicking.");
                        }
                    }

                    if (shouldAutoClickTop)
                    {
                        Helpers.Logger.LogInfo($"ExecuteActionAsync: AUTO-CLICKING candidate '{candidates[0].Label}' via {autoClickReason}.");
                        var point = Helpers.VisualTargetClickPointResolver.Resolve(candidates[0], visualIdentify.TargetPhrase);
                        if (point == System.Drawing.Point.Empty)
                        {
                            point = candidates[0].Center;
                            Helpers.Logger.LogDebug($"ExecuteActionAsync: resolver returned empty, using center point.");
                        }
                        Helpers.Logger.LogInfo($"ExecuteActionAsync: click point: ({point.X},{point.Y}) for '{candidates[0].Label}'.");
                        PerformLeftClickAtPoint(point);
                        Helpers.Logger.LogInfo($"ExecuteActionAsync: PerformLeftClickAtPoint completed.");

                        VisualCandidateOverlayForm.HideOverlay();
                        Helpers.VisualCandidateSessionStore.Clear();
                        Helpers.Logger.LogInfo($"Visual identify auto-clicked top candidate via {autoClickReason}.");
                        return $"Clicked visual target '{candidates[0].Label}' (confidence {candidates[0].Confidence:0.00}).";
                    }
                    else
                    {
                        Helpers.Logger.LogInfo($"ExecuteActionAsync: NOT auto-clicking - confidence too low. Top: {candidates[0].Confidence:0.00} < threshold: {workaroundMinConfidence:0.00}.");
                    }

                    // Reorder candidates for display (top-to-bottom, left-to-right) so numbering
                    // follows the visual layout. Persist this ordering in the session so
                    // spoken "choose N" maps to the displayed numbers.
                    var displayCandidates = candidates
                        .OrderBy(c => c.Bounds.Top)
                        .ThenBy(c => c.Bounds.Left)
                        .ToList();

                    var session2 = new Models.VisualTargetSession
                    {
                        Query = visualIdentify.TargetPhrase,
                        CreatedUtc = DateTime.UtcNow,
                        Candidates = displayCandidates
                    };
                    Helpers.VisualCandidateSessionStore.SetSession(session2);

                    VisualCandidateOverlayForm.ShowCandidates(displayCandidates, AppSettings.Instance.VisualTargeting.OverlayTimeoutMs);
                    return $"Found {displayCandidates.Count} visual candidates for '{visualIdentify.TargetPhrase}'. Say 'choose 1' to 'choose {displayCandidates.Count}'.";
                }
                catch (Exception ex)
                {
                    Helpers.Logger.LogError($"VisualIdentifyClickAction failed: {ex.Message}");
                    return $"Visual identify failed: {ex.Message}";
                }
            }
            else if (action is VisualShowCandidatesAction)
            {
                var session = Helpers.VisualCandidateSessionStore.GetSession();
                if (session == null || session.Candidates.Count == 0)
                {
                    return "No active visual candidate session.";
                }

                VisualCandidateOverlayForm.ShowCandidates(session.Candidates, AppSettings.Instance.VisualTargeting.OverlayTimeoutMs);
                return $"Showing {session.Candidates.Count} visual candidates for '{session.Query}'.";
            }
            else if (action is VisualChooseCandidateAction chooseVisual)
            {
                if (!Helpers.VisualCandidateSessionStore.TryGetCandidate(chooseVisual.CandidateNumber, out var candidate) || candidate == null)
                {
                    return $"Candidate {chooseVisual.CandidateNumber} is not available. Say 'show candidates' to view options again.";
                }

                VisualCandidateOverlayForm.HideOverlay();
                System.Threading.Thread.Sleep(30);

                var session = Helpers.VisualCandidateSessionStore.GetSession();
                var point = Helpers.VisualTargetClickPointResolver.Resolve(candidate, session?.Query ?? string.Empty);
                if (point == System.Drawing.Point.Empty)
                {
                    point = candidate.Center;
                }
                Helpers.Logger.LogInfo($"Visual choose click point: ({point.X},{point.Y}) for '{candidate.Label}'.");
                PerformLeftClickAtPoint(point);

                Helpers.VisualCandidateSessionStore.Clear();
                return $"Clicked candidate {chooseVisual.CandidateNumber}: {candidate.Label}.";
            }
            // Quick Clicks execution handlers removed (feature deprecated).
            else if (action is AdjustAutoClickDelayAction adj)
            {
                AppendLog($"[DEBUG] ExecuteActionAsync: Executing AdjustAutoClickDelayAction (change={adj.Change})\n");
                if (string.Equals(adj.Change, "faster", StringComparison.OrdinalIgnoreCase))
                {
                    return NaturalCommands.Helpers.AutoClickManager.Faster();
                }
                else if (string.Equals(adj.Change, "slower", StringComparison.OrdinalIgnoreCase))
                {
                    return NaturalCommands.Helpers.AutoClickManager.Slower();
                }
                else
                {
                    return "Unknown auto-click adjustment command.";
                }
            }
            else if (action is SetWindowAlwaysOnTopAction setTop)
            {
                // This feature was disabled because it caused accidental "always on top" states.
                AppendLog("[INFO] ExecuteActionAsync: Attempt to set always-on-top blocked by configuration.\n");
                return "Setting windows always-on-top has been disabled.";
            }
            // Focus window by title substring
            else if (action is FocusWindowAction focusAction)
            {
                // Workaround for speech-to-text errors: normalize common misrecognitions
                string normalized = focusAction.WindowTitleSubstring
                    .Replace("tall window", "tool window", StringComparison.OrdinalIgnoreCase)
                    .Replace("tall", "tool", StringComparison.OrdinalIgnoreCase)
                    .Replace("propertie", "properties", StringComparison.OrdinalIgnoreCase)
                    .Replace("property's", "properties", StringComparison.OrdinalIgnoreCase)
                    .Replace("propertys", "properties", StringComparison.OrdinalIgnoreCase)
                    .Replace("'s", "s", StringComparison.OrdinalIgnoreCase)
                    .Replace("'", "", StringComparison.OrdinalIgnoreCase)
                    .Trim();
                if (!string.Equals(normalized, focusAction.WindowTitleSubstring, StringComparison.OrdinalIgnoreCase))
                {
                    AppendLog($"[INFO] Normalized window title substring from '{focusAction.WindowTitleSubstring}' to '{normalized}'\n");
                }
                string? procName = CurrentApplicationHelper.GetCurrentProcessName();
                if (procName == "devenv" && (normalized.Contains("tool window") || normalized.Contains("tool")))
                {
                    // Map common tool window names to Visual Studio commands
                    var toolWindowMap = new System.Collections.Generic.Dictionary<string, string>(System.StringComparer.OrdinalIgnoreCase)
                    {
                        { "properties", "View.PropertiesWindow" },
                        { "solution explorer", "View.SolutionExplorer" },
                        { "toolbox", "View.Toolbox" },
                        { "error list", "View.ErrorList" },
                        { "output", "View.Output" },
                        { "class view", "View.ClassView" },
                        { "command window", "View.CommandWindow" },
                        { "task list", "View.TaskList" },
                        { "team explorer", "View.TeamExplorer" },
                        { "bookmarks", "View.BookmarkWindow" },
                        { "call hierarchy", "View.CallHierarchy" },
                        { "object browser", "View.ObjectBrowser" },
                        { "find results", "View.FindResults1" },
                        { "pending changes", "View.PendingChanges" },
                        { "git changes", "View.GitChanges" },
                        { "git repository", "View.GitRepository" },
                        { "breakpoints", "Debug.BreakpointsWindow" },
                        { "locals", "Debug.Locals" },
                        { "autos", "Debug.Autos" },
                        { "watch", "Debug.Watch" },
                        { "immediate", "Debug.Immediate" },
                        { "call stack", "Debug.CallStack" },
                        { "threads", "Debug.Threads" },
                        { "modules", "Debug.Modules" },
                        { "memory", "Debug.Memory1" },
                        { "disassembly", "Debug.Disassembly" },
                        { "output window", "View.Output" }
                    };
                    // Try to find the tool window name in the normalized string
                    foreach (var kvp in toolWindowMap)
                    {
                        if (normalized.Contains(kvp.Key))
                        {
                            bool vsResult = NaturalCommands.Helpers.VisualStudioHelper.ExecuteCommand(kvp.Value);
                            if (vsResult)
                            {
                                AppendLog($"[INFO] Focused Visual Studio tool window: '{kvp.Key}' via command '{kvp.Value}'\n");
                                return $"Focused Visual Studio tool window: {kvp.Key}";
                            }
                            else
                            {
                                AppendLog($"[WARN] Failed to focus Visual Studio tool window: '{kvp.Key}' via command '{kvp.Value}'\n");
                                return $"Failed to focus Visual Studio tool window: {kvp.Key}";
                            }
                        }
                    }
                    // If no known tool window matched, log and fall through to window title focus
                    AppendLog($"[WARN] No known Visual Studio tool window matched in '{normalized}', falling back to window title focus.\n");
                }
                // Try the normalized string first (for non-VS or unknown tool windows)
                bool focused = NaturalCommands.Helpers.WindowFocusHelper.FocusWindowByTitle(normalized);
                // Fallback: if the normalized string contains 'properties' and 'tool window', also try just 'properties'
                if (!focused && normalized.Contains("properties", StringComparison.OrdinalIgnoreCase) && normalized.Contains("tool window", StringComparison.OrdinalIgnoreCase))
                {
                    AppendLog($"[INFO] Fallback: trying to focus window with title containing just 'properties'\n");
                    focused = NaturalCommands.Helpers.WindowFocusHelper.FocusWindowByTitle("properties");
                }
                if (focused)
                {
                    AppendLog($"[INFO] Focused window with title containing: '{normalized}' (or fallback)\n");
                    return $"Focused window: {normalized}";
                }
                else
                {
                    AppendLog($"[WARN] Could not find window with title containing: '{normalized}' (or fallback)\n");
                    return $"Could not find window with title containing: {normalized}";
                }
            }
            // Multi-action sequences: run a list of actions in order
            else if (action is NaturalCommands.RunMultipleActionsAction multiAction)
            {
                AppendLog($"[DEBUG] ExecuteActionAsync: Running multi-action '{multiAction.Name}' with {multiAction.Actions?.Count ?? 0} steps\n");
                string lastResult = string.Empty;
                if (multiAction.Actions != null)
                {
                    foreach (var sub in multiAction.Actions)
                    {
                        try
                        {
                            lastResult = ExecuteActionAsync(sub);
                            AppendLog($"[DEBUG] ExecuteActionAsync: multi-step result: {lastResult}\n");
                        }
                        catch (Exception ex)
                        {
                            AppendLog($"[ERROR] ExecuteActionAsync: multi-step failed: {ex.Message}\n");
                            if (!multiAction.ContinueOnError)
                            {
                                return $"Multi-action '{multiAction.Name}' aborted: {ex.Message}";
                            }
                        }
                        try { System.Threading.Thread.Sleep(Math.Max(0, multiAction.DelayMsBetween)); } catch { }
                    }
                }
                return $"Executed multi-action '{multiAction.Name}'" + (string.IsNullOrEmpty(lastResult) ? "." : $": {lastResult}");
            }
            else if (action is OpenFolderAction folder)
            {
                string path;
                switch (folder.KnownFolder.ToLower())
                {
                    case "downloads":
                        path = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
                        break;
                    case "documents":
                        path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                        break;
                    case "mycomputer":
                    case "computer":
                        // Use SH: MyComputerFolder so Explorer shows drives and available space
                        path = "shell:MyComputerFolder";
                        break;
                    default:
                        path = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                        break;
                }

                // Try to focus existing Explorer window for Downloads
                bool focused = FocusExistingExplorerWindow(path);
                if (focused)
                {
                    AppendLog("Focused existing Downloads window\n");
                    return $"Focused existing window: {folder.KnownFolder} ({path})";
                }
                // Otherwise, open new window
                var psi = new System.Diagnostics.ProcessStartInfo("explorer.exe", path)
                {
                    UseShellExecute = true
                };
                System.Diagnostics.Process.Start(psi);
                if (folder.KnownFolder.Equals("Downloads", StringComparison.OrdinalIgnoreCase))
                {
                    AppendLog("Opened folder: Downloads\n");
                }
                return $"Opened folder: {folder.KnownFolder} ({path})";
            // No closing brace here; keep method open for further action handlers
            }
            else if (action is ShowHelpAction)
            {
                string helpText = "Available commands:\n" +
                    "- Move this window to the left/right\n" +
                    "- Maximize this window\n" +
                    "- Open downloads/documents/This PC\n" +
                    "- Move window to other screen\n" +
                    "- (More natural commands can be added)";
                // If the command list was already shown in a dialog, skip the tray notification
                if (_suppressNextHelpNotification)
                {
                    _suppressNextHelpNotification = false;
                    return helpText;
                }
                TrayNotificationHelper.ShowNotification("NaturalCommands.NET Help", helpText, 7000);
                return helpText;
            }
            else if (action is LaunchAppAction app)
            {
                return NaturalCommands.Helpers.AppLauncher.Launch(app);
            }
            else if (action is NaturalCommands.OpenWebsiteAction website)
            {
                var result = WebsiteNavigator.LaunchWebsite(website.Url);
                AppendLog(result + "\n");
                return result;
            }
            else if (action is NaturalCommands.SendKeysAction keys)
            {
                return NaturalCommands.Helpers.KeySender.SendKeys(keys);
            }
            else if (action is NaturalCommands.EmojiAction emojiAction)
            {
                return NaturalCommands.Helpers.EmojiService.TypeEmoji(emojiAction);
            }
            else if (action is NaturalCommands.ExecuteVSCommandAction vsCmd)
            {
                return NaturalCommands.Helpers.VSCommandHandler.ExecuteVSCommand(vsCmd);
            }
            else if (action is NaturalCommands.ShowLettersAction showLetters)
            {
                try
                {
                    UIElementOverlayForm.ShowOverlay(showLetters.ScopeToActiveWindow);
                    AppendLog("[INFO] Show letters overlay displayed.\n");
                    return "Show letters overlay displayed. Type letters to click elements, or press ESC to cancel.";
                }
                catch (Exception ex)
                {
                    AppendLog($"[ERROR] Failed to show letters overlay: {ex.Message}\n");
                    return $"Failed to show letters overlay: {ex.Message}";
                }
            }
            else if (action is NaturalCommands.ShowTaskbarAction)
            {
                try
                {
                    // Prefer enumerating the Taskbar HWND directly so we get the middle app area reliably
                    IntPtr trayHwnd = NaturalCommands.Helpers.WindowFocusHelper.GetTaskbarHwnd();
                    bool focused = NaturalCommands.Helpers.WindowFocusHelper.FocusTaskbar();
                    if (trayHwnd != IntPtr.Zero)
                    {
                        AppendLog($"[DEBUG] Found Taskbar hwnd: {trayHwnd}\n");
                        UIElementOverlayForm.ShowOverlayForWindow(trayHwnd);
                        AppendLog("[INFO] Show letters overlay displayed for Taskbar.\n");
                        return "Taskbar focused and show letters overlay displayed. Type letters to click elements, or press ESC to cancel.";
                    }
                    else
                    {
                        if (!focused) AppendLog("[WARN] FocusTaskbar returned false; proceeding to default overlay anyway.\n");
                        UIElementOverlayForm.ShowOverlay(true);
                        AppendLog("[INFO] Show letters overlay displayed.\n");
                        return "Show letters overlay displayed. Type letters to click elements, or press ESC to cancel.";
                    }
                }
                catch (Exception ex)
                {
                    AppendLog($"[ERROR] Failed to show taskbar overlay: {ex.Message}\n");
                    return $"Failed to show letters overlay for taskbar: {ex.Message}";
                }
            }
            else if (action is NaturalCommands.ShowDesktopAction)
            {
                try
                {
                    // First ensure the desktop is shown and focused (Win+D via FocusDesktop)
                    bool focused = NaturalCommands.Helpers.WindowFocusHelper.FocusDesktop();
                    // Re-fetch the desktop hwnd after attempting to focus it
                    IntPtr desktopHwnd = NaturalCommands.Helpers.WindowFocusHelper.GetDesktopHwnd();
                    if (desktopHwnd != IntPtr.Zero)
                    {
                        AppendLog($"[DEBUG] Found Desktop hwnd: {desktopHwnd}\n");
                        UIElementOverlayForm.ShowOverlayForWindow(desktopHwnd);
                        AppendLog("[INFO] Show letters overlay displayed for Desktop.\n");
                        return "Desktop focused and show letters overlay displayed. Type letters to activate icons, or press ESC to cancel.";
                    }
                    else
                    {
                        if (!focused) AppendLog("[WARN] FocusDesktop returned false; proceeding to default overlay anyway.\n");
                        UIElementOverlayForm.ShowOverlay(true);
                        AppendLog("[INFO] Show letters overlay displayed.\n");
                        return "Show letters overlay displayed. Type letters to click elements, or press ESC to cancel.";
                    }
                }
                catch (Exception ex)
                {
                    AppendLog($"[ERROR] Failed to show desktop overlay: {ex.Message}\n");
                    return $"Failed to show letters overlay for desktop: {ex.Message}";
                }
            }
            else
            {
                // Always log 'No matching action' for unknown action types
                AppendLog("No matching action\n");
                return "Unknown action type.";
            }
        }

        public string HandleNaturalAsync(string text)
        {
            AppendLog($"[DEBUG] Entered HandleNaturalAsync with text: '{text}'\n");
            string lowerText = text.ToLowerInvariant();
            AppendLog($"[DEBUG] lowerText: '{lowerText}', contains 'focus': {lowerText.Contains("focus")}, contains 'zoom': {lowerText.Contains("zoom")}\n");
            // Intercept 'focus zoom' and similar before AI fallback
            if (lowerText.Contains("focus") && lowerText.Contains("zoom"))
            {
                AppendLog("[DEBUG] Intercepted 'focus zoom' before AI fallback.\n");
                bool focused = NaturalCommands.Helpers.WindowFocusHelper.FocusZoom();
                if (focused)
                {
                    AppendLog("[DEBUG] Successfully focused Zoom window.\n");
                    return "[Natural mode] Focused Zoom window.";
                }
                else
                {
                    AppendLog("[DEBUG] Zoom window not found, showing Ctrl+Alt+Tab fallback.\n");
                    var sim = new WindowsInput.InputSimulator();
                    sim.Keyboard.ModifiedKeyStroke(
                        new[] {
                            WindowsInput.Native.VirtualKeyCode.CONTROL,
                            WindowsInput.Native.VirtualKeyCode.MENU
                        },
                        WindowsInput.Native.VirtualKeyCode.TAB);
                    return "[Natural mode] Sent Ctrl+Alt+Tab for app switcher (focus fallback)";
                }
            }
            var actionTask = InterpretAsync(text);
            actionTask.Wait();
            var action = actionTask.Result;
            string actionTypeName = action != null ? action.GetType().Name : "null";
            AppendLog($"[DEBUG] HandleNaturalAsync: Action type: {actionTypeName}\n");
            if (action == null)
            {
                // Try normalization/fuzzy correction before AI fallback
                string normalized = text
                    .Replace("tall window", "tool window", StringComparison.OrdinalIgnoreCase)
                    .Replace("tall", "tool", StringComparison.OrdinalIgnoreCase)
                    .Replace("propertie", "properties", StringComparison.OrdinalIgnoreCase)
                    .Replace("property's", "properties", StringComparison.OrdinalIgnoreCase)
                    .Replace("propertys", "properties", StringComparison.OrdinalIgnoreCase)
                    .Replace("'s", "s", StringComparison.OrdinalIgnoreCase)
                    .Replace("'", "", StringComparison.OrdinalIgnoreCase)
                    .Trim();
                if (!string.Equals(normalized, text, StringComparison.OrdinalIgnoreCase))
                {
                    AppendLog($"[INFO] HandleNaturalAsync: Normalized input from '{text}' to '{normalized}' before AI fallback\n");
                    var retryTask = InterpretAsync(normalized);
                    retryTask.Wait();
                    var retryAction = retryTask.Result;
                    if (retryAction != null)
                    {
                        AppendLog($"[INFO] HandleNaturalAsync: Normalized retry succeeded with action: {retryAction.GetType().Name}\n");
                        action = retryAction;
                        actionTypeName = action.GetType().Name;
                    }
                }
            }
            if (action == null)
            {
                // Robust fallback for 'focus windows terminal' and 'focus' commands
                if (lowerText.Contains("focus") && lowerText.Contains("windows terminal"))
                {
                    AppendLog("[DEBUG] Attempting to focus Windows Terminal window (HandleNaturalAsync single-arg)\n");
                    bool focused = NaturalCommands.Helpers.WindowFocusHelper.FocusWindowsTerminal();
                    if (focused)
                    {
                        AppendLog("[DEBUG] Successfully focused Windows Terminal window.\n");
                        return "[Natural mode] Focused Windows Terminal window.";
                    }
                    else
                    {
                        AppendLog("[DEBUG] Windows Terminal not found, showing Ctrl+Alt+Tab fallback.\n");
                        var sim = new WindowsInput.InputSimulator();
                        sim.Keyboard.ModifiedKeyStroke(
                            new[] {
                                WindowsInput.Native.VirtualKeyCode.CONTROL,
                                WindowsInput.Native.VirtualKeyCode.MENU
                            },
                            WindowsInput.Native.VirtualKeyCode.TAB);
                        return "[Natural mode] Sent Ctrl+Alt+Tab for app switcher (focus fallback)";
                    }
                }
                if (lowerText.Contains("focus"))
                {
                    AppendLog("[DEBUG] Focus fallback: sending Ctrl+Alt+Tab (HandleNaturalAsync single-arg)\n");
                    var sim = new WindowsInput.InputSimulator();
                    sim.Keyboard.ModifiedKeyStroke(
                        new[] {
                            WindowsInput.Native.VirtualKeyCode.CONTROL,
                            WindowsInput.Native.VirtualKeyCode.MENU
                        },
                        WindowsInput.Native.VirtualKeyCode.TAB);
                    return "[Natural mode] Sent Ctrl+Alt+Tab for app switcher (focus fallback)";
                }
                
                // First, try Talon fallback with AI semantic assistance (if enabled)
                AppendLog($"[DEBUG] HandleNaturalAsync: Checking Talon fallback first. EnableFallback={AppSettings.Instance.Talon.EnableFallback}\n");
                if (AppSettings.Instance.Talon.EnableFallback)
                {
                    try
                    {
                        AppendLog($"[DEBUG] HandleNaturalAsync: Attempting Talon resolution for: '{text}'\n");
                        var talonActionTask = NaturalCommands.Helpers.TalonCommandRouter.ResolveActionAsync(text);
                        talonActionTask.Wait();
                        var talonAction = talonActionTask.Result;
                        AppendLog($"[DEBUG] HandleNaturalAsync: Talon ResolveActionAsync returned: {(talonAction == null ? "null" : $"'{talonAction.TalonCommand}'")}\n");
                        if (talonAction != null)
                        {
                            AppendLog($"[DEBUG] HandleNaturalAsync: Talon fallback matched '{talonAction.TalonCommand}' ({talonAction.MatchSource})\n");
                            var talonResult = ExecuteActionAsync(talonAction);
                            return $"[Natural mode] {talonResult}";
                        }
                        else
                        {
                            AppendLog("[DEBUG] HandleNaturalAsync: Talon returned null (no match)\n");
                        }
                    }
                    catch (Exception ex)
                    {
                        AppendLog($"[WARN] HandleNaturalAsync: Talon fallback error: {ex.Message}\n");
                    }
                }

                // If Talon could not resolve, fallback to OpenAI for general command interpretation
                AppendLog($"[DEBUG] HandleNaturalAsync: Fallback to OpenAI for: {text}\n");
                string? currentApp = NaturalCommands.CurrentApplicationHelper.GetCurrentProcessName();
                string aiInput = text;
                if (!string.IsNullOrWhiteSpace(currentApp))
                {
                    aiInput += $"\nCurrentApplication: {currentApp}";
                }

                var aiActionTask = InterpretWithAIAsync(aiInput);
                aiActionTask.Wait();
                var aiAction = aiActionTask.Result;
                AppendLog($"[DEBUG] HandleNaturalAsync: OpenAI Action type: {(aiAction == null ? "null" : aiAction.GetType().Name)}\n");

                if (aiActionTask.IsCompletedSuccessfully && aiAction != null)
                {
                    // If the original text was 'close tab', override AI fallback to always send Ctrl+W
                    if (text.Trim().Equals("close tab", StringComparison.InvariantCultureIgnoreCase) || text.Trim().Equals("closed tab", StringComparison.InvariantCultureIgnoreCase))
                    {
                        var closeTabAction = new CloseTabAction();
                        AppendLog($"[DEBUG] Overriding AI fallback for 'close tab' to always send Ctrl+W.\n");
                        var resultOverride = ExecuteActionAsync(closeTabAction);
                        return $"[Natural mode] {resultOverride}";
                    }
                    // Otherwise, use AI result as normal
                    AppendLog($"[DEBUG] HandleNaturalAsync: See '[AI] Raw response' above for actual AI output.\n");
                    var aiResult = ExecuteActionAsync(aiAction);
                    return $"[Natural mode] {aiResult}";
                }
                else
                {
                    AppendLog("[DEBUG] HandleNaturalAsync: Talon fallback is disabled\n");
                }

                AppendLog("No matching action\n");
                // Show auto-closing message box for unmatched command
                AppendLog($"[DEBUG] About to show AutoClosingMessageBox for: {text}\n");
                int timeoutMs = 5000;
                int timeoutSec = timeoutMs / 1000;
                string msg = $"No matching action for: {text}\n(This will close in {timeoutSec} seconds)";
                NaturalCommands.AutoClosingMessageBox.Show(msg, "Command Not Recognized", timeoutMs);
                return $"[Natural mode] No matching action for: {text}";
            }
            var result = ExecuteActionAsync(action);
            return $"[Natural mode] {result}";
        }

        public string HandleNaturalAsync(string text, List<(string Command, string Description)> availableCommands)
        {
            var actionTask = InterpretAsync(text, availableCommands);
            actionTask.Wait();
            var action = actionTask.Result;
            NaturalCommands.Helpers.Logger.LogDebug($"HandleNaturalAsync: Action type: {(action == null ? "null" : action.GetType().Name)}");
            if (action == null)
            {
                // If the command contains 'focus windows terminal', try direct focus
                if (text.Contains("focus") && text.Contains("windows terminal"))
                {
                    AppendLog("[DEBUG] Attempting to focus Windows Terminal window (HandleNaturalAsync)\n");
                    bool focused = NaturalCommands.Helpers.WindowFocusHelper.FocusWindowsTerminal();
                    if (focused)
                    {
                        AppendLog("[DEBUG] Successfully focused Windows Terminal window.\n");
                        return "[Natural mode] Focused Windows Terminal window.";
                    }
                    else
                        if (action == null)
                        {
                            // Robust fallback for Windows Terminal focus
                            string lowerText = text.ToLowerInvariant();
                            if (lowerText.Contains("focus") && lowerText.Contains("windows terminal"))
                            {
                                AppendLog("[DEBUG] Attempting to focus Windows Terminal window (HandleNaturalAsync overload)\n");
                                bool focusedInner = NaturalCommands.Helpers.WindowFocusHelper.FocusWindowsTerminal();
                                if (focusedInner)
                                {
                                    AppendLog("[DEBUG] Successfully focused Windows Terminal window.\n");
                                    return "[Natural mode] Focused Windows Terminal window.";
                                }
                                else
                                {
                                    AppendLog("[DEBUG] Windows Terminal not found, showing Ctrl+Alt+Tab fallback.\n");
                                    var sim = new WindowsInput.InputSimulator();
                                    sim.Keyboard.ModifiedKeyStroke(
                                        new[] {
                                            WindowsInput.Native.VirtualKeyCode.CONTROL,
                                            WindowsInput.Native.VirtualKeyCode.MENU
                                        },
                                        WindowsInput.Native.VirtualKeyCode.TAB);
                                    return "[Natural mode] Sent Ctrl+Alt+Tab for app switcher (focus fallback)";
                                }
                            }
                            // Fallback for any 'focus' command
                            if (lowerText.Contains("focus"))
                            {
                                NaturalCommands.Helpers.Logger.LogDebug("Focus fallback: sending Ctrl+Alt+Tab (HandleNaturalAsync overload)");
                                var sim = new WindowsInput.InputSimulator();
                                sim.Keyboard.ModifiedKeyStroke(
                                    new[] {
                                        WindowsInput.Native.VirtualKeyCode.CONTROL,
                                        WindowsInput.Native.VirtualKeyCode.MENU
                                    },
                                    WindowsInput.Native.VirtualKeyCode.TAB);
                                return "[Natural mode] Sent Ctrl+Alt+Tab for app switcher (focus fallback)";
                            }
                            // ...existing code...
                        }
                }
                // If the command contains 'focus', trigger Ctrl+Alt+Tab fallback
                if (text.Contains("focus"))
                {
                    NaturalCommands.Helpers.Logger.LogDebug("Focus fallback: sending Ctrl+Alt+Tab (HandleNaturalAsync)");
                    var sim = new WindowsInput.InputSimulator();
                    sim.Keyboard.ModifiedKeyStroke(
                        new[] {
                            WindowsInput.Native.VirtualKeyCode.CONTROL,
                            WindowsInput.Native.VirtualKeyCode.MENU
                        },
                        WindowsInput.Native.VirtualKeyCode.TAB);
                    return "[Natural mode] Sent Ctrl+Alt+Tab for app switcher (focus fallback)";
                }
                // Suggest available commands if no match
                var suggestions = string.Join(", ", availableCommands.Select(c => c.Command));
                return $"[Natural mode] No matching action for: {text}. Available commands: {suggestions}";
            }
            var result = ExecuteActionAsync(action);
            return $"[Natural mode] {result}";
        }

        public System.Threading.Tasks.Task<ActionBase?> InterpretAsync(string text, List<(string Command, string Description)> availableCommands)
        {
            text = (text ?? string.Empty).ToLowerInvariant().Trim();
            // Remove polite modifiers
            text = RemovePoliteModifiers(text);
            text = WordReplacementLoader.Apply(text);
            NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync input: {text}");

            // Fuzzy match against available commands
            var bestMatch = availableCommands
                .Select(cmd => (cmd.Command, Score: GetSimilarityScore(text, cmd.Command)))
                .OrderByDescending(x => x.Score)
                .FirstOrDefault();

            // Threshold for fuzzy match (can be tuned)
            if (bestMatch.Score > 0.6)
            {
                // Map best match to action
                switch (bestMatch.Command)
                {
                    case "maximize window":
                        var action = new MoveWindowAction(Target: "active", Monitor: "current", Position: "center", WidthPercent: 100, HeightPercent: 100);
                        NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync matched: {action.GetType().Name}");
                        return System.Threading.Tasks.Task.FromResult<ActionBase?>(action);
                    case "move window to left half":
                        var leftAction = new MoveWindowAction(Target: "active", Monitor: "current", Position: "left", WidthPercent: 50, HeightPercent: 100);
                        NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync matched: {leftAction.GetType().Name} (left half)");
                        return System.Threading.Tasks.Task.FromResult<ActionBase?>(leftAction);
                    case "move window to right half":
                        var rightAction = new MoveWindowAction(Target: "active", Monitor: "current", Position: "right", WidthPercent: 50, HeightPercent: 100);
                        NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync matched: {rightAction.GetType().Name} (right half)");
                        return System.Threading.Tasks.Task.FromResult<ActionBase?>(rightAction);
                    case "move window to other monitor":
                        var nextAction = new MoveWindowAction(Target: "active", Monitor: "next", Position: "", WidthPercent: 0, HeightPercent: 0);
                        NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync matched: {nextAction.GetType().Name} (next monitor)");
                        return System.Threading.Tasks.Task.FromResult<ActionBase?>(nextAction);
                    
                    case "open downloads":
                        var downloadsAction = new OpenFolderAction("Downloads");
                        NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync matched: {downloadsAction.GetType().Name} (downloads)");
                        return System.Threading.Tasks.Task.FromResult<ActionBase?>(downloadsAction);
                    case "open documents":
                        var documentsAction = new OpenFolderAction("Documents");
                        NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync matched: {documentsAction.GetType().Name} (documents)");
                        return System.Threading.Tasks.Task.FromResult<ActionBase?>(documentsAction);
                    case "close tab":
                        var closeTabAction = new CloseTabAction();
                        NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync matched: {closeTabAction.GetType().Name} (close tab)");
                        return System.Threading.Tasks.Task.FromResult<ActionBase?>(closeTabAction);
                    case "send keys":
                        var sendKeysAction = new SendKeysAction(text.Replace("press ", ""));
                        NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync matched: {sendKeysAction.GetType().Name} (send keys)");
                        return System.Threading.Tasks.Task.FromResult<ActionBase?>(sendKeysAction);
                    case "launch app":
                        var launchAppAction = new LaunchAppAction(text.Replace("open ", ""));
                        NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync matched: {launchAppAction.GetType().Name} (launch app)");
                        return System.Threading.Tasks.Task.FromResult<ActionBase?>(launchAppAction);
                    case "focus app":
                        // Not implemented, stub
                        break;
                    case "show help":
                        var helpAction = new ShowHelpAction();
                        NaturalCommands.Helpers.Logger.LogDebug($"InterpretAsync matched: {helpAction.GetType().Name} (help)");
                        return System.Threading.Tasks.Task.FromResult<ActionBase?>(helpAction);
                }
            }
            NaturalCommands.Helpers.Logger.LogDebug("InterpretAsync: No match, returning null");
            return System.Threading.Tasks.Task.FromResult<ActionBase?>(null);
        }

        // Simple similarity score (normalized longest common subsequence)
        private static double GetSimilarityScore(string input, string command)
        {
            input = input.ToLowerInvariant();
            command = command.ToLowerInvariant();
            int lcs = LongestCommonSubsequence(input, command);
            return (double)lcs / Math.Max(input.Length, command.Length);
        }

        // Longest common subsequence algorithm
        private static int LongestCommonSubsequence(string a, string b)
        {
            int[,] dp = new int[a.Length + 1, b.Length + 1];
            for (int i = 1; i <= a.Length; i++)
            {
                for (int j = 1; j <= b.Length; j++)
                {
                    if (a[i - 1] == b[j - 1])
                        dp[i, j] = dp[i - 1, j - 1] + 1;
                    else
                        dp[i, j] = Math.Max(dp[i - 1, j], dp[i, j - 1]);
                }
            }
            return dp[a.Length, b.Length];
        }

        // Helper to enumerate all monitor handles
        private static IEnumerable<IntPtr> GetAllMonitors()
        {
            var monitors = new List<IntPtr>();
            bool MonitorEnum(IntPtr hMonitor, IntPtr hdcMonitor, ref RECT lprcMonitor, IntPtr dwData)
            {
                monitors.Add(hMonitor);
                return true;
            }
            EnumDisplayMonitors(IntPtr.Zero, IntPtr.Zero, MonitorEnum, IntPtr.Zero);
            return monitors;
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool EnumDisplayMonitors(IntPtr hdc, IntPtr lprcClip, MonitorEnumProc lpfnEnum, IntPtr dwData);
        private delegate bool MonitorEnumProc(IntPtr hMonitor, IntPtr hdcMonitor, ref RECT lprcMonitor, IntPtr dwData);
                [System.Runtime.InteropServices.DllImport("user32.dll")]
                private static extern bool SetForegroundWindow(IntPtr hWnd);
            // Helper: Focus existing Explorer window for a given path
            private static bool FocusExistingExplorerWindow(string folderPath)
            {
                try
                {
                    // Use COM to enumerate Explorer windows
                    Type? shellWindowsType = Type.GetTypeFromProgID("Shell.Application");
                    if (shellWindowsType == null)
                        return false;
                    dynamic? shellWindows = Activator.CreateInstance(shellWindowsType);
                    if (shellWindows == null)
                        return false;
                    foreach (var window in shellWindows.Windows())
                    {
                        string url = "";
                        try { url = window.LocationURL as string ?? ""; } catch { }
                        string hwndStr = "";
                        try { hwndStr = window.HWND.ToString(); } catch { }
                        // Convert file:///C:/Users/.../Downloads to local path
                        if (url.StartsWith("file:///", StringComparison.OrdinalIgnoreCase))
                        {
                            string winPath = Uri.UnescapeDataString(url.Substring(8).Replace('/', '\\'));
                            // Compare normalized paths
                            if (string.Equals(System.IO.Path.GetFullPath(winPath), System.IO.Path.GetFullPath(folderPath), StringComparison.OrdinalIgnoreCase))
                            {
                                // Focus window
                                IntPtr hWnd = IntPtr.Zero;
                                if (IntPtr.TryParse(hwndStr, out hWnd) && hWnd != IntPtr.Zero)
                                {
                                    SetForegroundWindow(hWnd);
                                    return true;
                                }
                                // Fallback: try window.HWND as int
                                try
                                {
                                    hWnd = (IntPtr)window.HWND;
                                    if (hWnd != IntPtr.Zero)
                                    {
                                        SetForegroundWindow(hWnd);
                                        return true;
                                    }
                                }
                                catch { }
                            }
                        }
                    }
                }
                catch { }
                return false;
            }
    }
}
