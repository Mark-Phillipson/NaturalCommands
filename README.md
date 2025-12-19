## Talon Voice Integration

NaturalCommands can be triggered using Talon Voice by saying **"natural"** (or **"nat"**) followed by the command you want to execute. For example:

- `natural what can I say`
- `natural dictate`

### Talon File Commands

```talon
^(natural | nat) <user.text>$:
	user.run_application_csharp_natural(text)
^(natural | nat) dictate$:
	speech.disable()
	user.run_application_csharp_natural("dictate")
```

### Python Code

```python
def run_application_csharp_natural(naturalCommand: str):
	"""runs the natural command with the given text"""
	import os
	commandline = r'C:\Users\MPhil\source\repos\NaturalCommands\bin\Release\net10.0-windows\NaturalCommands.exe'
	args1 = ' ' + r'/natural' + ' '
	args2 = '' + r'/' + naturalCommand + ''
	arguments = [args1, args2]
	cwd = os.path.dirname(commandline)
	print(commandline)
	print(naturalCommand)
	ui.launch(path=commandline, args=arguments, cwd=cwd)
```
# NaturalCommands

Execute Windows and Visual Studio commands using natural language.

## Overview

NaturalCommands is a lightweight Windows application that maps natural language phrases to system and Visual Studio actions. It includes helpers for sending keys, managing windows, running processes, and integrating with AI-based interpreters.

## Features

- Map natural language to keyboard and window actions
- Voice dictation helpers and multi-action support
- Visual Studio command helpers and shortcuts
- **Show letters** feature for voice-based UI element navigation

## Voice Command: "What can I say?"

You can say **"what can I say"** at any time to display a list of available commands.

## Prerequisites

- Windows 10 or later
- [.NET SDK 10.0+](https://dotnet.microsoft.com/en-us/download)

## Build & Run

```bash
# Restore, build and run
dotnet build NaturalCommands.csproj  
dotnet build NaturalCommands.csproj  -c Release
```

## Listen mode (resident hotkey)

Run the app in resident mode to open the voice dictation UI from anywhere via a global hotkey.

- Start resident mode:
	- `dotnet run --framework net10.0-windows -- listen`
	- or run the built exe: `NaturalCommands.exe listen`
- Hotkey: **Win+Ctrl+H**
- Fallback: use the system tray menu item **Open Voice Dictation (Win+Ctrl+H)**

When voice typing stops, the **Send Command** button is focused and the form’s **Enter** key submits (AcceptButton), so you can press Enter to send the captured text.

## Show Letters Feature 🎯

The "show letters" feature enables voice-based navigation of UI elements by overlaying letter labels on clickable items.

### Usage

1. Say **"show letters"** or **"natural show letters"**
2. A transparent overlay will appear showing two-letter labels (a, b, c, ..., aa, ab, ...) on all clickable elements
3. Type the letters corresponding to the element you want to click using the Talon alphabet
4. The element will be clicked automatically
5. Press **ESC** to cancel without clicking

### Examples

```bash
# Show letter labels on clickable elements
dotnet run --framework net10.0-windows -- natural "show letters"

# Then type letters like "a", "b", "ab", etc. to click the corresponding element
```

### Supported Elements

The feature works with:
- Buttons
- Links
- Menu items
- Checkboxes
- Radio buttons
- Tab items
- Combo boxes

### Works With

- **Web browsers** (Edge, Chrome, Firefox)
- **Desktop applications** (Windows native apps)
- Any application that supports Windows UI Automation API

## Contributing

Contributions are welcome — please open issues and pull requests on GitHub.

## Steam detection (experimental) 🎮

- The app can now detect installed Steam games by scanning your Steam libraries (appmanifest_*.acf) and will allow launching games by voice using commands like `play <game name>`.
- Launches use the Steam URI (`steam://rungameid/<appid>`) so Steam handles the actual game start.
- Currently Steam is the only supported source; detection is experimental — open an issue if you encounter games that are not detected or mismatched.

## License

This project is licensed under the MIT License — see the `LICENSE` file for details.
