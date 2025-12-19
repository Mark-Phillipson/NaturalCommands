## Build the Application
```pwsh
# Build the main project (requires .NET 10)
dotnet clean NaturalCommands.csproj 

dotnet build NaturalCommands.csproj  
dotnet build NaturalCommands.csproj  -c Release
dotnet run --framework net10.0-windows -- listen
dotnet publish ./NaturalCommands.csproj -c Release -f net10.0-windows -r win-x64 `
  --self-contained true -p:PublishSingleFile=true `
  -o "./bin/Release/net10.0-windows/win-x64/publish"

## Publish and register Startup (one command)
You can publish and create a per-user Startup shortcut with the included script (no admin required for the Startup shortcut):

```pwsh
# Publish self-contained and add a Startup shortcut that runs the app with `-- listen`
powershell -ExecutionPolicy Bypass -File .\scripts\publish-and-register-startup.ps1 -SelfContained -CreateStartupShortcut
```

To publish framework-dependent (smaller) and skip the shortcut:

```pwsh
powershell -ExecutionPolicy Bypass -File .\scripts\publish-and-register-startup.ps1 -NoSelfContained -CreateStartupShortcut:$false
```
```

## Output Files

- Application log: [app.log](bin/bin/app.log)
- Latest AI prompt file: [latest_ai_prompt.md](bin/bin/latest_ai_prompt.md) (overwritten each run; contains the system prompt and the latest user input)
 

## Example Natural Language Actions
```pwsh
# Run with a specific action (replace with your desired command)
dotnet run --framework net10.0-windows -- natural "open calculator"
dotnet run --framework net10.0-windows -- natural "type hello world"
dotnet run --framework net10.0-windows -- natural "open notepad"
dotnet run --framework net10.0-windows -- natural "move this window to the other screen"
dotnet run --framework net10.0-windows -- natural "maximize this window"
dotnet run --framework net10.0-windows -- natural "put this window on the left half"
dotnet run --framework net10.0-windows -- natural "put this window on the right half"
dotnet run --framework net10.0-windows -- natural " close this window"
dotnet run --framework net10.0-windows -- natural "natural dictate"  # Opens the voice dictation form for speaking or typing natural language commands
dotnet run --framework net10.0-windows -- natural "show letters"  # Display letter labels on clickable UI elements for voice navigation
```

## Listen mode (resident hotkey)

Start the app in resident mode to show a tray icon and register a global hotkey:

```pwsh
dotnet run --framework net10.0-windows -- listen
```

- Hotkey: **Win+Ctrl+H** (opens the existing voice dictation form)
- If the hotkey is already used by another app, you can still open dictation from the tray menu.

While dictation is active, the form shows a “Listening… — press Enter to send” hint. When dictation stops, **Send Command** is focused and pressing **Enter** submits.

## Window Management Actions
You can use natural language to control window position and size:

- Move to next monitor: `dotnet run --framework net10.0-windows -- natural "move this window to the other screen"`
- Maximize: `dotnet run --framework net10.0-windows -- natural "maximize this window"`
- Left half: `dotnet run --framework net10.0-windows -- natural "put this window on the left half"`
- Right half: `dotnet run --framework net10.0-windows -- natural "put this window on the right half"`


## Running Integration Tests (Not viable do not use)
To run all integration tests (including window management):

```pwsh
dotnet test Tests/IntegrationTests.csproj
```

Tests will verify that window management, app launching, folder opening, and other actions work as expected.

## Notes

## Visual Studio Natural Language Command Tests

You can run these commands from the integrated terminal inside Visual Studio:

| Natural Language Command         | Expected Visual Studio Action         | Canonical Command Name      |
|----------------------------------|--------------------------------------|----------------------------|
| build the solution               | Build the entire solution            | Build.BuildSolution        |
| build the project                | Build the current project            | Build.BuildProject         |
| start debugging                  | Start debugging the startup project  | Debug.Start                |
| start application                | Start without debugging              | Debug.StartWithoutDebugging|
| stop debugging                   | Stop debugging                       | Debug.StopDebugging        |
| close tab                        | Close the current document tab       | Window.CloseDocumentWindow |
| format document                  | Format the current document          | Edit.FormatDocument        |
| find in files                    | Open the Find in Files dialog        | Edit.FindinFiles           |
| go to definition                 | Go to definition of symbol           | Edit.GoToDefinition        |
| rename symbol                    | Rename the selected symbol           | Refactor.Rename            |
| show solution explorer           | Focus Solution Explorer              | View.SolutionExplorer      |
| open recent files                | Show recent files                    | File.RecentFiles           |

Example usage:
```pwsh
dotnet run --framework net10.0-windows -- natural "build the solution"
dotnet run --framework net10.0-windows -- natural "start debugging"
dotnet run --framework net10.0-windows -- natural "close tab"
dotnet run --framework net10.0-windows -- natural "format document"
```

