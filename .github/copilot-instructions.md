# Copilot Instructions for NaturalCommands

## Build and test commands

Use `NaturalCommands_dotnet.sln` as the default solution entry point; both solution files currently include the same projects.

```pwsh
# Build the app and test project
dotnet build .\NaturalCommands_dotnet.sln -c Debug

# Build the app for release output
dotnet build .\NaturalCommands.csproj -c Release

# Run the xUnit suite
dotnet test .\NaturalCommands_dotnet.sln -c Debug --no-build

# Run a single xUnit test
dotnet test .\Tests\NaturalCommands_NET.Tests.csproj -c Debug --no-build --filter "FullyQualifiedName~NaturalCommands_NET.Tests.VisualTargetingServiceTests.MatchesTokens_BasicAndFuzzy"

# CLI smoke tests that validate log-driven behavior
powershell -ExecutionPolicy Bypass -File .\Tests\cli\open-microsoft-paint.test.ps1
powershell -ExecutionPolicy Bypass -File .\Tests\cli\focus-properties-tool-window.test.ps1

# Run the app in its main modes
dotnet run --project .\NaturalCommands.csproj --framework net10.0-windows -- natural "show letters"
dotnet run --project .\NaturalCommands.csproj --framework net10.0-windows -- listen
```

## High-level architecture

- `Program.cs` is the mode dispatcher. It handles `natural`, `listen`, `listen-notifications`, `ticker`, `ticker-file`, `ticker-test`, and `export-vs-commands`.
- `Commands.HandleNaturalAsync()` is only a thin wrapper; the real command pipeline lives in `NaturalLanguageInterpreter`.
- `NaturalLanguageInterpreter.InterpretAsync()` does deterministic parsing first: normalize input, apply `word_replacements.json`, match built-in commands, multi-actions, Visual Studio / Windows Terminal / Explorer shortcuts, quick-click commands, voice dictation, show-letters, and visual targeting. `HandleNaturalAsync()` only falls back to Talon matching and then OpenAI if deterministic parsing returns no action.
- Action shapes are centralized in `ActionModels.cs`. Execution is centralized in `NaturalLanguageInterpreter.ExecuteActionAsync(...)`, which is where new action records are wired into actual behavior.
- Resident mode is implemented by `ListenModeApplicationContext`. It owns the tray icon, the `Win+Ctrl+H` hotkey, Quick Clicks menu actions, and the file watchers used for resident-process coordination.
- Cross-process coordination is file-based under `%LOCALAPPDATA%\NaturalCommands`: listen mode watches `.quick_clicks_command`, and ticker messages are delivered through `.ticker_payload`.
- Visual targeting is a layered subsystem: `VisualTargetingService` tries local UI Automation first, can escalate to cloud vision when enabled, and can fall back to local OCR before `VisualCandidateOverlayForm` presents numbered choices.

## Key conventions

- `NaturalLanguageInterpreter.cs` explicitly says not to add new methods there. Prefer a new helper or focused class, then call it from the interpreter.
- When adding a new command, follow the existing three-part pattern: add or reuse an `ActionBase` record in `ActionModels.cs`, teach `InterpretAsync()` how to return it, and handle execution in `ExecuteActionAsync(...)` or a dedicated helper.
- Keep logging on `NaturalCommands.Helpers.Logger`; do not introduce `Console.WriteLine` as the primary diagnostics path. The log file is reset per process start and can be redirected with `NATURALCOMMANDS_LOG`.
- Runtime configuration is path-sensitive. `settings.json`, `multi_actions.json`, `vs_commands.json`, `emoji_mappings.json`, Talon cache files, and other runtime assets are resolved relative to `AppDomain.CurrentDomain.BaseDirectory`, so new config files usually need matching `CopyToOutputDirectory` handling in `NaturalCommands.csproj`.
- `AppSettings.Instance` is the shared settings entry point. Most feature flags and persistent behavior hang off that singleton.
- Overlay-style features need a WinForms UI context. `Program.cs` has a special path that initializes WinForms and keeps a message pump alive for auto-click, Quick Clicks, and visual-targeting commands; do not bypass that when adding UI-backed actions.
- `QuickClickOverlayForm` and its resident-mode menu wiring exist, but `Helpers.QuickClickLoader` is currently stubbed out and returns no profiles. Do not assume `quick_clicks.json` persistence is active unless you re-enable that loader.
