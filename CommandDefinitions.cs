using System;
using System.Collections.Generic;

namespace NaturalCommands
{
    public static class CommandDefinitions
    {
        public static readonly List<(string Command, string Description)> AvailableCommands = new()
        {
            ("maximize window", "Maximize the active window"),
            ("move window to left half", "Move the active window to the left half of the screen"),
            ("move window to right half", "Move the active window to the right half of the screen"),
            ("move window to other monitor", "Move the active window to the next monitor"),
            ("open downloads", "Open the Downloads folder"),
            ("open documents", "Open the Documents folder"),
            ("open settings", "Open the application settings"),
            ("open my computer", "Open This PC / show drives and available space"),
            ("open this pc", "Open This PC / show drives and available space"),
            ("close tab", "Close the current tab in supported applications"),
            ("send keys", "Send a key sequence to the active window"),
            ("launch app", "Launch a specified application"),
            ("focus app", "Focus a specified application window"),
            ("show help", "Show help and available commands"),
            ("emoji set <name> <emoji>", "Set an emoji for a named shortcut (e.g. emoji set happy 😀)"),
            ("emoji <name>", "Insert the configured emoji for the given name"),
            ("emoji <emoji>", "Insert the given emoji immediately")
        };
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
        public static readonly Dictionary<string, string> VisualStudioCanonicalMappings = new(StringComparer.OrdinalIgnoreCase)
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
            { "format document", "Edit.FormatDocument" },
            { "find in files", "Edit.FindinFiles" },
            { "go to definition", "Edit.GoToDefinition" },
            { "rename symbol", "Refactor.Rename" },
            { "show solution explorer", "View.SolutionExplorer" },
            { "open recent files", "File.RecentFiles" }
        };
        public static readonly Dictionary<string, ActionBase> PopularCommands = new(StringComparer.OrdinalIgnoreCase)
        {
            { "debug application", new ExecuteVSCommandAction("Debug.Start") },
            { "run application", new ExecuteVSCommandAction("Debug.StartWithoutDebugging") },
            { "stop application", new ExecuteVSCommandAction("Debug.StopDebugging") }
        };

        // Windows Terminal specific commands - maps natural language to keyboard shortcuts
        public static readonly List<(string Command, string Description)> WindowsTerminalCommands = new()
        {
            ("new tab", "Open a new tab (Ctrl+Shift+T)"),
            ("close tab", "Close the current tab (Ctrl+Shift+W)"),
            ("close pane", "Close the current pane (Ctrl+Shift+W)"),
            ("split horizontal", "Split pane horizontally (Alt+Shift+-)"),
            ("split vertical", "Split pane vertically (Alt+Shift++)"),
            ("split down", "Split pane down (Alt+Shift+-)"),
            ("split right", "Split pane right (Alt+Shift++)"),
            ("duplicate pane", "Duplicate the current pane (Alt+Shift+D)"),
            ("next tab", "Switch to the next tab (Ctrl+Tab)"),
            ("previous tab", "Switch to the previous tab (Ctrl+Shift+Tab)"),
            ("next pane", "Move focus to the next pane (Alt+Down/Right)"),
            ("previous pane", "Move focus to the previous pane (Alt+Up/Left)"),
            ("focus pane up", "Move focus to the pane above (Alt+Up)"),
            ("focus pane down", "Move focus to the pane below (Alt+Down)"),
            ("focus pane left", "Move focus to the pane on the left (Alt+Left)"),
            ("focus pane right", "Move focus to the pane on the right (Alt+Right)"),
            ("resize pane up", "Resize pane up (Alt+Shift+Up)"),
            ("resize pane down", "Resize pane down (Alt+Shift+Down)"),
            ("resize pane left", "Resize pane left (Alt+Shift+Left)"),
            ("resize pane right", "Resize pane right (Alt+Shift+Right)"),
            ("open settings", "Open terminal settings (Ctrl+,)"),
            ("open command palette", "Open command palette (Ctrl+Shift+P)"),
            ("toggle fullscreen", "Toggle full screen mode (F11)"),
            ("toggle focus mode", "Toggle focus mode (Ctrl+Shift+Enter)"),
            ("toggle always on top", "Toggle always on top (Alt+Enter)"),
            ("find", "Open find dialog (Ctrl+Shift+F)"),
            ("scroll up", "Scroll up (Ctrl+Shift+Up)"),
            ("scroll down", "Scroll down (Ctrl+Shift+Down)"),
            ("scroll page up", "Scroll up one page (Ctrl+Shift+PageUp)"),
            ("scroll page down", "Scroll down one page (Ctrl+Shift+PageDown)"),
            ("zoom in", "Increase font size (Ctrl++)"),
            ("zoom out", "Decrease font size (Ctrl+-)"),
            ("reset zoom", "Reset font size (Ctrl+0)"),
            ("copy", "Copy selected text (Ctrl+Shift+C)"),
            ("paste", "Paste from clipboard (Ctrl+Shift+V)"),
            ("rename tab", "Rename the current tab (Ctrl+Shift+R)"),
            ("switch to tab 1", "Switch to tab 1 (Ctrl+Alt+1)"),
            ("switch to tab 2", "Switch to tab 2 (Ctrl+Alt+2)"),
            ("switch to tab 3", "Switch to tab 3 (Ctrl+Alt+3)"),
            ("switch to tab 4", "Switch to tab 4 (Ctrl+Alt+4)"),
            ("switch to tab 5", "Switch to tab 5 (Ctrl+Alt+5)"),
        };

        // Maps natural language commands to Windows Terminal shortcuts
        public static readonly Dictionary<string, string> WindowsTerminalShortcuts = new(StringComparer.OrdinalIgnoreCase)
        {
            // Tab management
            { "new tab", "ctrl+shift+t" },
            { "open new tab", "ctrl+shift+t" },
            { "create tab", "ctrl+shift+t" },
            { "add tab", "ctrl+shift+t" },
            { "close tab", "ctrl+shift+w" },
            { "close pane", "ctrl+shift+w" },
            { "next tab", "ctrl+tab" },
            { "previous tab", "ctrl+shift+tab" },
            { "tab right", "ctrl+tab" },
            { "tab left", "ctrl+shift+tab" },

            // Pane splitting
            { "split horizontal", "alt+shift+minus" },
            { "split horizontally", "alt+shift+minus" },
            { "split down", "alt+shift+minus" },
            { "split vertical", "alt+shift+plus" },
            { "split vertically", "alt+shift+plus" },
            { "split right", "alt+shift+plus" },
            { "duplicate pane", "alt+shift+d" },
            { "duplicate", "alt+shift+d" },

            // Pane focus navigation
            { "focus pane up", "alt+up" },
            { "focus up", "alt+up" },
            { "pane up", "alt+up" },
            { "focus pane down", "alt+down" },
            { "focus down", "alt+down" },
            { "pane down", "alt+down" },
            { "focus pane left", "alt+left" },
            { "focus left", "alt+left" },
            { "pane left", "alt+left" },
            { "focus pane right", "alt+right" },
            { "focus right", "alt+right" },
            { "pane right", "alt+right" },
            { "next pane", "alt+right" },
            { "previous pane", "alt+left" },

            // Pane resizing
            { "resize pane up", "alt+shift+up" },
            { "resize up", "alt+shift+up" },
            { "resize pane down", "alt+shift+down" },
            { "resize down", "alt+shift+down" },
            { "resize pane left", "alt+shift+left" },
            { "resize left", "alt+shift+left" },
            { "resize pane right", "alt+shift+right" },
            { "resize right", "alt+shift+right" },

            // Settings and commands
            { "open settings", "ctrl+," },
            { "settings", "ctrl+," },
            { "terminal settings", "ctrl+," },
            { "open command palette", "ctrl+shift+p" },
            { "command palette", "ctrl+shift+p" },
            { "commands", "ctrl+shift+p" },

            // View modes
            { "toggle fullscreen", "f11" },
            { "fullscreen", "f11" },
            { "full screen", "f11" },
            { "toggle focus mode", "ctrl+shift+enter" },
            { "focus mode", "ctrl+shift+enter" },
            { "toggle always on top", "alt+enter" },

            // Find
            { "find", "ctrl+shift+f" },
            { "search", "ctrl+shift+f" },
            { "find text", "ctrl+shift+f" },

            // Scrolling
            { "scroll up", "ctrl+shift+up" },
            { "scroll down", "ctrl+shift+down" },
            { "scroll page up", "ctrl+shift+pageup" },
            { "scroll page down", "ctrl+shift+pagedown" },
            { "page up", "ctrl+shift+pageup" },
            { "page down", "ctrl+shift+pagedown" },

            // Zoom
            { "zoom in", "ctrl+plus" },
            { "zoom out", "ctrl+minus" },
            { "reset zoom", "ctrl+0" },
            { "increase font", "ctrl+plus" },
            { "decrease font", "ctrl+minus" },
            { "reset font", "ctrl+0" },

            // Copy/Paste
            { "copy", "ctrl+shift+c" },
            { "paste", "ctrl+shift+v" },

            // Tab rename
            { "rename tab", "ctrl+shift+r" },
            { "rename", "ctrl+shift+r" },

            // Switch to specific tabs
            { "switch to tab 1", "ctrl+alt+1" },
            { "tab 1", "ctrl+alt+1" },
            { "switch to tab 2", "ctrl+alt+2" },
            { "tab 2", "ctrl+alt+2" },
            { "switch to tab 3", "ctrl+alt+3" },
            { "tab 3", "ctrl+alt+3" },
            { "switch to tab 4", "ctrl+alt+4" },
            { "tab 4", "ctrl+alt+4" },
            { "switch to tab 5", "ctrl+alt+5" },
            { "tab 5", "ctrl+alt+5" },
            { "switch to tab 6", "ctrl+alt+6" },
            { "tab 6", "ctrl+alt+6" },
            { "switch to tab 7", "ctrl+alt+7" },
            { "tab 7", "ctrl+alt+7" },
            { "switch to tab 8", "ctrl+alt+8" },
            { "tab 8", "ctrl+alt+8" },
            { "switch to tab 9", "ctrl+alt+9" },
            { "tab 9", "ctrl+alt+9" },

            // Additional useful commands
            { "clear", "ctrl+shift+k" },
            { "clear terminal", "ctrl+shift+k" },
            { "clear screen", "ctrl+shift+k" },
            { "new window", "ctrl+shift+n" },
            { "open new window", "ctrl+shift+n" },
        };
    }
}
