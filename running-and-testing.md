## Build the Application
```pwsh
# Build the main project (requires .NET 10)
dotnet clean NaturalCommands.csproj 

dotnet build NaturalCommands.csproj  
dotnet build NaturalCommands.csproj  
dotnet build NaturalCommands.csproj  -c Release
dotnet run --framework net10.0-windows -c Release

dotnet publish ./NaturalCommands.csproj -c Release -f net10.0-windows -r win-x64
  --self-contained true -p:PublishSingleFile=true `
  -o "./bin/Release/net10.0-windowsdotnet clean NaturalCommands.csproj/win-x64/publish"
```
## Publish and No Startup (one command)
You can publish and create a per-user Startup shortcut with the included script (no admin required for the Startup shortcut):

```pwsh
# Publish self-contained and add a Startup shortcut that runs the app (no extra args required)
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
dotnet run --framework net10.0-windows -- natural "optical"
dotnet run --framework net10.0-windows -- natural "type hello world"
dotnet run --framework net10.0-windows -- natural "open notepad"
dotnet run --framework net10.0-windows -- natural "move this window to the other screen"
dotnet run --framework net10.0-windows -- natural "maximize this window"
dotnet run --framework net10.0-windows -- natural "put this window on the left half"
dotnet run --framework net10.0-windows -- natural "put this window on the right half"
dotnet run --framework net10.0-windows -- natural " close this window"
dotnet run --framework net10.0-windows -- natural "natural dictate"  # Opens the voice dictation form for speaking or typing natural language commands
dotnet run --framework net10.0-windows -- natural "show letters"  # Display letter labels on clickable UI elements for voice navigation
dotnet run --framework net10.0-windows -- natural "show taskbar"  # Focus the Taskbar then display letters for taskbar elements
```

-## Resident mode (tray hotkey)

Start the app to show a tray icon and register a global hotkey. By default the app will create the tray/menu at startup when `Settings -> Notifications -> ShowTrayOnStartup` is enabled.

```pwsh
dotnet run --framework net10.0-windows
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

Target for a Windows Shortcut:
```
"C:\Users\MPhil\source\repos\NaturalCommands\bin\Release\net10.0-windows\NaturalCommands.exe"
```

## Notifications

- The resident notification listener uses the WinRT `UserNotificationListener` API which requires user permission and (in many cases) package identity. If the listener cannot access the notification manager you may see errors such as `RequestAccess failed: ... CO_E_OBJNOTCONNECTED` or occasional `AppInfo` cast failures.
- Quick troubleshooting:
  - Run the app (no args) (`dotnet run --framework net10.0-windows`) and check Windows Settings > System > Notifications to see if `NaturalCommands` appears and can be enabled.
  - For robust support, grant package identity to the app (use a Packaging Project with "Package with External Location" or create a lightweight identity package). See the Microsoft docs: https://learn.microsoft.com/windows/apps/desktop/modernize/grant-identity-to-nonpackaged-apps-overview
- Note: The code was updated to handle these failures more gracefully and to apply exponential backoff for automatic restarts — rebuild the app to pick up the changes.
