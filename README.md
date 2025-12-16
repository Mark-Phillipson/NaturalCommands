# NaturalCommands

Execute Windows and Visual Studio commands using natural language.

## Overview

NaturalCommands is a lightweight Windows application that maps natural language phrases to system and Visual Studio actions. It includes helpers for sending keys, managing windows, running processes, and integrating with AI-based interpreters.

## Features

- Map natural language to keyboard and window actions
- Voice dictation helpers and multi-action support
- Visual Studio command helpers and shortcuts
- Unit tests and Playwright tests included

## Prerequisites

- Windows 10 or later
- [.NET SDK 10.0+](https://dotnet.microsoft.com/en-us/download)

## Build & Run

```bash
# Restore, build and run
dotnet build NaturalCommands.csproj  
dotnet build NaturalCommands.csproj  -c Release
```

## Contributing

Contributions are welcome — please open issues and pull requests on GitHub.

## Steam detection (experimental) 🎮

- The app can now detect installed Steam games by scanning your Steam libraries (appmanifest_*.acf) and will allow launching games by voice using commands like `play <game name>`.
- Launches use the Steam URI (`steam://rungameid/<appid>`) so Steam handles the actual game start.
- Currently Steam is the only supported source; detection is experimental — open an issue if you encounter games that are not detected or mismatched.

## License

This project is licensed under the MIT License — see the `LICENSE` file for details.
