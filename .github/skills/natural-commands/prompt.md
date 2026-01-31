# NaturalCommands Skill — Build & Run (Copilot)

This skill provides exact commands and examples to build, publish, and run the NaturalCommands application, and to start it in resident "listen" mode. Use these when a user asks how to build, publish, run, or use listen mode.

## Quick commands (PowerShell, Windows)
- Build and clean:
  - `dotnet clean NaturalCommands.csproj`
  - `dotnet build NaturalCommands.csproj`
  - `dotnet build NaturalCommands.csproj -c Release`

- Run in listen mode (resident hotkey / tray icon):
  - `dotnet run --framework net10.0-windows -- listen`

- Run a single natural command (examples):
  - `dotnet run --framework net10.0-windows -- natural "open calculator"`
  - `dotnet run --framework net10.0-windows -- natural "type hello world"`

- Publish (self-contained) & create Startup shortcut (script):
  - `powershell -ExecutionPolicy Bypass -File .\scripts\publish-and-register-startup.ps1 -SelfContained -CreateStartupShortcut`
  - Framework-dependent publish (no shortcut):
    `powershell -ExecutionPolicy Bypass -File .\scripts\publish-and-register-startup.ps1 -NoSelfContained -CreateStartupShortcut:$false`

## Notes & verification
- Application log: `bin/bin/app.log`. When validating behavior, check this log file for expected messages.
- The project has a policy not to run automated tests for verification; prefer manual log inspection.
- Examples for Visual Studio/IDE natural-language actions are available in `running-and-testing.md`.

## Guidance for Copilot responses
When users ask about building or running the project, provide:
- Exact shell commands they can copy-and-paste (PowerShell / Windows) for build, run, publish, listen mode.
- Short explanation of what each command does and any important flags (e.g., `-c Release`, `--framework net10.0-windows`).
- A reminder to check `bin/bin/app.log` for verification and to avoid running the test suite for verification unless explicitly requested.
- If asked about packaging/publishing, show the `powershell` script examples and state the output location.

If a user asks how to add a Startup shortcut or publish, include the `publish-and-register-startup.ps1` usage shown here and warn about `-SelfContained` vs framework-dependent outputs.

---
*Source of commands: `running-and-testing.md` in the repository.*
