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
            ("shutdown", "Stop/interrupt the current running process (Ctrl+C)"),
            ("stop", "Stop/interrupt the current running process (Ctrl+C)"),
            ("stop process", "Stop/interrupt the current running process (Ctrl+C)"),
            ("kill process", "Kill/interrupt the current running process (Ctrl+C)"),
            ("cancel", "Cancel/interrupt the current running process (Ctrl+C)"),
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

            // Process control - Ctrl+C to interrupt/stop running process
            { "shutdown", "ctrl+c" },
            { "stop", "ctrl+c" },
            { "stop process", "ctrl+c" },
            { "kill", "ctrl+c" },
            { "kill process", "ctrl+c" },
            { "terminate", "ctrl+c" },
            { "cancel", "ctrl+c" },
            { "interrupt", "ctrl+c" },
            { "abort", "ctrl+c" },
            { "break", "ctrl+c" },
            { "stop server", "ctrl+c" },
            { "stop application", "ctrl+c" },
            { "stop running", "ctrl+c" },
        };

        // .NET CLI commands that can be run in Windows Terminal
        public static readonly List<(string Command, string TerminalCommand, string Description)> DotNetCommands = new()
        {
            // Build commands
            ("dotnet build", "dotnet build", "Build a .NET project"),
            ("build", "dotnet build", "Build a .NET project"),
            ("build project", "dotnet build", "Build a .NET project"),
            ("build solution", "dotnet build", "Build a .NET project/solution"),
            ("dotnet build release", "dotnet build -c Release", "Build in Release configuration"),
            ("build release", "dotnet build -c Release", "Build in Release configuration"),
            
            // Clean commands
            ("dotnet clean", "dotnet clean", "Clean build outputs of a .NET project"),
            ("clean", "dotnet clean", "Clean build outputs of a .NET project"),
            ("clean project", "dotnet clean", "Clean build outputs of a .NET project"),
            ("clean solution", "dotnet clean", "Clean build outputs of a .NET project/solution"),
            
            // Run commands
            ("dotnet run", "dotnet run", "Build and run a .NET project"),
            ("run", "dotnet run", "Build and run a .NET project"),
            ("run project", "dotnet run", "Build and run a .NET project"),
            ("run application", "dotnet run", "Build and run a .NET project"),
            ("dotnet run release", "dotnet run -c Release", "Run in Release configuration"),
            ("run release", "dotnet run -c Release", "Run in Release configuration"),
            
            // Test commands
            ("dotnet test", "dotnet test", "Run unit tests"),
            ("test", "dotnet test", "Run unit tests"),
            ("run tests", "dotnet test", "Run unit tests"),
            ("run unit tests", "dotnet test", "Run unit tests"),
            ("dotnet test verbose", "dotnet test -v n", "Run unit tests with normal verbosity"),
            ("test verbose", "dotnet test -v n", "Run unit tests with normal verbosity"),
            
            // Restore commands
            ("dotnet restore", "dotnet restore", "Restore dependencies"),
            ("restore", "dotnet restore", "Restore dependencies"),
            ("restore packages", "dotnet restore", "Restore NuGet packages"),
            ("restore dependencies", "dotnet restore", "Restore dependencies"),
            
            // Publish commands
            ("dotnet publish", "dotnet publish", "Publish a .NET project for deployment"),
            ("publish", "dotnet publish", "Publish a .NET project for deployment"),
            ("publish project", "dotnet publish", "Publish a .NET project for deployment"),
            ("dotnet publish release", "dotnet publish -c Release", "Publish in Release configuration"),
            ("publish release", "dotnet publish -c Release", "Publish in Release configuration"),
            
            // Pack commands
            ("dotnet pack", "dotnet pack", "Create a NuGet package"),
            ("pack", "dotnet pack", "Create a NuGet package"),
            ("create package", "dotnet pack", "Create a NuGet package"),
            ("dotnet pack release", "dotnet pack -c Release", "Create a NuGet package in Release configuration"),
            ("pack release", "dotnet pack -c Release", "Create a NuGet package in Release configuration"),
            
            // Watch commands
            ("dotnet watch", "dotnet watch", "Start a file watcher that runs a command when files change"),
            ("watch", "dotnet watch", "Start a file watcher that runs a command when files change"),
            ("dotnet watch run", "dotnet watch run", "Watch and run on file changes"),
            ("watch run", "dotnet watch run", "Watch and run on file changes"),
            ("dotnet watch test", "dotnet watch test", "Watch and run tests on file changes"),
            ("watch test", "dotnet watch test", "Watch and run tests on file changes"),
            
            // Format commands
            ("dotnet format", "dotnet format", "Apply style preferences to a project or solution"),
            ("format", "dotnet format", "Apply style preferences to a project or solution"),
            ("format code", "dotnet format", "Apply style preferences to a project or solution"),
            
            // New project commands
            ("dotnet new console", "dotnet new console", "Create a new console application"),
            ("new console", "dotnet new console", "Create a new console application"),
            ("dotnet new classlib", "dotnet new classlib", "Create a new class library"),
            ("new class library", "dotnet new classlib", "Create a new class library"),
            ("dotnet new webapi", "dotnet new webapi", "Create a new Web API project"),
            ("new web api", "dotnet new webapi", "Create a new Web API project"),
            ("dotnet new mvc", "dotnet new mvc", "Create a new MVC web application"),
            ("new mvc", "dotnet new mvc", "Create a new MVC web application"),
            ("dotnet new blazor", "dotnet new blazor", "Create a new Blazor web application"),
            ("new blazor", "dotnet new blazor", "Create a new Blazor web application"),
            ("dotnet new list", "dotnet new list", "List available project templates"),
            ("new list", "dotnet new list", "List available project templates"),
            ("list templates", "dotnet new list", "List available project templates"),
            
            // Solution commands
            ("dotnet solution list", "dotnet sln list", "List projects in a solution"),
            ("solution list", "dotnet sln list", "List projects in a solution"),
            ("list projects", "dotnet sln list", "List projects in a solution"),
            
            // Package commands
            ("dotnet add package", "dotnet add package ", "Add a NuGet package (prompts for name)"),
            ("add package", "dotnet add package ", "Add a NuGet package (prompts for name)"),
            ("dotnet list package", "dotnet list package", "List NuGet packages"),
            ("list packages", "dotnet list package", "List NuGet packages"),
            ("list nuget packages", "dotnet list package", "List NuGet packages"),
            ("dotnet remove package", "dotnet remove package ", "Remove a NuGet package (prompts for name)"),
            ("remove package", "dotnet remove package ", "Remove a NuGet package (prompts for name)"),
            
            // Tool commands
            ("dotnet tool list", "dotnet tool list", "List installed .NET tools"),
            ("list tools", "dotnet tool list", "List installed .NET tools"),
            ("dotnet tool install", "dotnet tool install ", "Install a .NET tool (prompts for name)"),
            ("install tool", "dotnet tool install ", "Install a .NET tool (prompts for name)"),
            ("dotnet tool update", "dotnet tool update ", "Update a .NET tool (prompts for name)"),
            ("update tool", "dotnet tool update ", "Update a .NET tool (prompts for name)"),
            
            // Info commands
            ("dotnet info", "dotnet --info", "Display .NET information"),
            ("dotnet version", "dotnet --version", "Display .NET SDK version"),
            ("dotnet list sdks", "dotnet --list-sdks", "Display installed SDKs"),
            ("list sdks", "dotnet --list-sdks", "Display installed SDKs"),
            ("dotnet list runtimes", "dotnet --list-runtimes", "Display installed runtimes"),
            ("list runtimes", "dotnet --list-runtimes", "Display installed runtimes"),
            
            // EF Core commands (common if EF tools installed)
            ("dotnet ef migrations add", "dotnet ef migrations add ", "Add a new EF Core migration (prompts for name)"),
            ("add migration", "dotnet ef migrations add ", "Add a new EF Core migration (prompts for name)"),
            ("dotnet ef migrations list", "dotnet ef migrations list", "List EF Core migrations"),
            ("list migrations", "dotnet ef migrations list", "List EF Core migrations"),
            ("dotnet ef database update", "dotnet ef database update", "Update database to latest migration"),
            ("update database", "dotnet ef database update", "Update database to latest migration"),
            ("dotnet ef database drop", "dotnet ef database drop", "Drop the database"),
            ("drop database", "dotnet ef database drop", "Drop the database"),
            
            // User secrets
            ("dotnet user-secrets init", "dotnet user-secrets init", "Initialize user secrets for the project"),
            ("init user secrets", "dotnet user-secrets init", "Initialize user secrets for the project"),
            ("dotnet user-secrets list", "dotnet user-secrets list", "List user secrets"),
            ("list user secrets", "dotnet user-secrets list", "List user secrets"),
            ("list secrets", "dotnet user-secrets list", "List user secrets"),
            
            // Dev certs
            ("dotnet dev-certs https trust", "dotnet dev-certs https --trust", "Trust the HTTPS development certificate"),
            ("trust dev cert", "dotnet dev-certs https --trust", "Trust the HTTPS development certificate"),
            ("trust https certificate", "dotnet dev-certs https --trust", "Trust the HTTPS development certificate"),
        };

        // Maps natural language to dotnet commands (for fuzzy matching)
        public static readonly Dictionary<string, (string TerminalCommand, string Description)> DotNetCommandMappings = new(StringComparer.OrdinalIgnoreCase)
        {
            // Build
            { "dotnet build", ("dotnet build", "Build a .NET project") },
            { "build", ("dotnet build", "Build a .NET project") },
            { "build project", ("dotnet build", "Build a .NET project") },
            { "build solution", ("dotnet build", "Build a .NET project/solution") },
            { "build the project", ("dotnet build", "Build a .NET project") },
            { "build the solution", ("dotnet build", "Build a .NET project/solution") },
            { "dotnet build release", ("dotnet build -c Release", "Build in Release configuration") },
            { "build release", ("dotnet build -c Release", "Build in Release configuration") },
            { "build in release", ("dotnet build -c Release", "Build in Release configuration") },
            
            // Clean
            { "dotnet clean", ("dotnet clean", "Clean build outputs") },
            { "clean", ("dotnet clean", "Clean build outputs") },
            { "clean project", ("dotnet clean", "Clean build outputs") },
            { "clean solution", ("dotnet clean", "Clean build outputs") },
            { "clean the project", ("dotnet clean", "Clean build outputs") },
            { "clean the solution", ("dotnet clean", "Clean build outputs") },
            
            // Run
            { "dotnet run", ("dotnet run", "Build and run a .NET project") },
            { "run", ("dotnet run", "Build and run a .NET project") },
            { "run project", ("dotnet run", "Build and run a .NET project") },
            { "run application", ("dotnet run", "Build and run a .NET project") },
            { "run the project", ("dotnet run", "Build and run a .NET project") },
            { "run the application", ("dotnet run", "Build and run a .NET project") },
            { "dotnet run release", ("dotnet run -c Release", "Run in Release configuration") },
            { "run release", ("dotnet run -c Release", "Run in Release configuration") },
            { "run in release", ("dotnet run -c Release", "Run in Release configuration") },
            
            // Test
            { "dotnet test", ("dotnet test", "Run unit tests") },
            { "test", ("dotnet test", "Run unit tests") },
            { "run tests", ("dotnet test", "Run unit tests") },
            { "run unit tests", ("dotnet test", "Run unit tests") },
            { "run the tests", ("dotnet test", "Run unit tests") },
            { "execute tests", ("dotnet test", "Run unit tests") },
            
            // Restore
            { "dotnet restore", ("dotnet restore", "Restore dependencies") },
            { "restore", ("dotnet restore", "Restore dependencies") },
            { "restore packages", ("dotnet restore", "Restore NuGet packages") },
            { "restore dependencies", ("dotnet restore", "Restore dependencies") },
            { "restore nuget", ("dotnet restore", "Restore NuGet packages") },
            
            // Publish
            { "dotnet publish", ("dotnet publish", "Publish for deployment") },
            { "publish", ("dotnet publish", "Publish for deployment") },
            { "publish project", ("dotnet publish", "Publish for deployment") },
            { "publish application", ("dotnet publish", "Publish for deployment") },
            
            // Pack
            { "dotnet pack", ("dotnet pack", "Create a NuGet package") },
            { "pack", ("dotnet pack", "Create a NuGet package") },
            { "create package", ("dotnet pack", "Create a NuGet package") },
            { "create nuget package", ("dotnet pack", "Create a NuGet package") },
            
            // Watch
            { "dotnet watch", ("dotnet watch", "Watch for file changes") },
            { "watch", ("dotnet watch", "Watch for file changes") },
            { "dotnet watch run", ("dotnet watch run", "Watch and run on changes") },
            { "watch run", ("dotnet watch run", "Watch and run on changes") },
            { "watch and run", ("dotnet watch run", "Watch and run on changes") },
            { "dotnet watch test", ("dotnet watch test", "Watch and test on changes") },
            { "watch test", ("dotnet watch test", "Watch and test on changes") },
            { "watch tests", ("dotnet watch test", "Watch and test on changes") },
            
            // Format
            { "dotnet format", ("dotnet format", "Format code") },
            { "format", ("dotnet format", "Format code") },
            { "format code", ("dotnet format", "Format code") },
            { "format the code", ("dotnet format", "Format code") },
            
            // Info
            { "dotnet info", ("dotnet --info", "Display .NET information") },
            { "dotnet version", ("dotnet --version", "Display .NET SDK version") },
            { "version", ("dotnet --version", "Display .NET SDK version") },
            { "dotnet list sdks", ("dotnet --list-sdks", "Display installed SDKs") },
            { "list sdks", ("dotnet --list-sdks", "Display installed SDKs") },
            { "dotnet list runtimes", ("dotnet --list-runtimes", "Display installed runtimes") },
            { "list runtimes", ("dotnet --list-runtimes", "Display installed runtimes") },
            
            // Common new project templates
            { "new console", ("dotnet new console", "Create a new console application") },
            { "new class library", ("dotnet new classlib", "Create a new class library") },
            { "new web api", ("dotnet new webapi", "Create a new Web API project") },
            { "new mvc", ("dotnet new mvc", "Create a new MVC web application") },
            { "new blazor", ("dotnet new blazor", "Create a new Blazor web application") },
            { "list templates", ("dotnet new list", "List available project templates") },
            
            // Solution and package commands
            { "solution list", ("dotnet sln list", "List projects in a solution") },
            { "list projects", ("dotnet sln list", "List projects in a solution") },
            { "list packages", ("dotnet list package", "List NuGet packages") },
            { "list nuget packages", ("dotnet list package", "List NuGet packages") },
            { "list tools", ("dotnet tool list", "List installed .NET tools") },
            
            // EF Core
            { "add migration", ("dotnet ef migrations add ", "Add a new EF Core migration") },
            { "list migrations", ("dotnet ef migrations list", "List EF Core migrations") },
            { "update database", ("dotnet ef database update", "Update database to latest migration") },
            { "drop database", ("dotnet ef database drop", "Drop the database") },
            
            // User secrets
            { "init user secrets", ("dotnet user-secrets init", "Initialize user secrets") },
            { "list user secrets", ("dotnet user-secrets list", "List user secrets") },
            { "list secrets", ("dotnet user-secrets list", "List user secrets") },
            
            // Dev certs
            { "trust dev cert", ("dotnet dev-certs https --trust", "Trust HTTPS dev certificate") },
            { "trust https certificate", ("dotnet dev-certs https --trust", "Trust HTTPS dev certificate") },
        };
    }
}
