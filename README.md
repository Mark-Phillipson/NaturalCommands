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

## Contributing

Contributions are welcome — please open issues and pull requests on GitHub.

## Steam detection (experimental) 🎮

- The app can now detect installed Steam games by scanning your Steam libraries (appmanifest_*.acf) and will allow launching games by voice using commands like `play <game name>`.
- Launches use the Steam URI (`steam://rungameid/<appid>`) so Steam handles the actual game start.
- Currently Steam is the only supported source; detection is experimental — open an issue if you encounter games that are not detected or mismatched.

## License

This project is licensed under the MIT License — see the `LICENSE` file for details.
