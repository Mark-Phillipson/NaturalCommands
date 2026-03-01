---
applyTo: '**'
---
# Copilot Instructions for NaturalCommands

## Testing & Verification
- **Do not run automated tests.** Verification is done via manual inspection of `bin/app.log`
- Use `Logger.LogDebug()`, `Logger.LogInfo()`, or `Logger.LogError()` for all diagnostic output
- The log path can be overridden via `NATURALCOMMANDS_LOG` environment variable

## Architecture Overview
This is a **.NET 10 Windows Forms/WPF hybrid** app that translates natural language voice commands into Windows automation actions.

### Core Data Flow
1. **Entry:** `Program.cs` parses CLI args (`natural <command>` or `listen` mode)
2. **Interpretation:** `NaturalLanguageInterpreter.cs` matches text → `ActionBase` (see `ActionModels.cs`)
3. **Execution:** Each action type has execution logic in `NaturalLanguageInterpreter` or `Helpers/` classes
4. **AI Fallback:** Unrecognized commands go to OpenAI API (requires `OPENAI_API_KEY` env var)

### Key Classes
| Class | Purpose |
|-------|---------|
| `ActionModels.cs` | All action types as C# records (e.g., `LaunchAppAction`, `SendKeysAction`) |
| `CommandDefinitions.cs` | Static mappings of natural phrases → actions |
| `NaturalLanguageInterpreter.cs` | Main command router (⚠️ large file, avoid adding methods here) |
| `ListenModeApplicationContext.cs` | System tray resident mode with `Win+Ctrl+H` hotkey |
| `VoiceDictationForm.cs` | Voice input UI for dictation mode |
| `UIElementOverlayForm.cs` | "Show letters" feature overlay for voice-based UI navigation |

### Helper Classes (`Helpers/`)
- `Logger.cs` – Central logging (writes to `bin/app.log`)
- `UIAutomationHelper.cs` – Windows UI Automation for "show letters" feature
- `WindowManager.cs`, `Win32ApiHelper.cs` – Window manipulation via Win32
- `VisualStudioHelper.cs`, `VSCommandHandler.cs` – VS automation via DTE
- `KeySender.cs` – Keyboard simulation via `InputSimulator`

## Build & Run
```bash
dotnet build NaturalCommands.csproj -c Release
# Run a command:
.\bin\Release\net10.0-windows\NaturalCommands.exe natural "maximize window"
# Resident listen mode:
.\bin\Release\net10.0-windows\NaturalCommands.exe listen
```

## Adding New Commands
1. Add action record to `ActionModels.cs`: `public record MyNewAction(...) : ActionBase;`
2. Add phrase mapping in `CommandDefinitions.cs` (if static mapping)
3. Add execution logic in `NaturalLanguageInterpreter.cs` (existing switch/match block)
4. For multi-step commands, add entry to `multi_actions.json`

## Configuration Files
| File | Purpose |
|------|---------|
| `settings.json` | Runtime settings (auto-click delay, hotkeys, etc.) |
| `multi_actions.json` | Composite command sequences |
| `vs_commands.json` | Visual Studio command mappings |
| `emoji_mappings.json` | Voice-activated emoji shortcuts |
| `word_replacements.json` | Speech recognition corrections |
| `openai_prompt.md` | System prompt for AI fallback interpretation |

## Conventions
- Use `NaturalCommands.Helpers.Logger` for all logging (never `Console.WriteLine`)
- Win32 imports go in `Helpers/Win32ApiHelper.cs` or `WindowUtils.cs`
- Settings are accessed via singleton `AppSettings.Instance`
- New features should be in separate files under `Helpers/` (NaturalLanguageInterpreter is too large)

## External Integrations
- **Talon Voice:** Triggered via `natural <command>` (see README for Talon config)
- **OpenAI API:** For fuzzy command matching when exact match fails
- **Visual Studio DTE:** For IDE automation commands

