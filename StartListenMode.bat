@echo off
REM Start NaturalCommands in listen mode with system tray icon
REM This keeps the application running to support features like auto-click

echo Starting NaturalCommands in listen mode...
echo.
echo System tray icon will appear in the notification area.
echo Right-click the icon to access:
echo   - Open Voice Dictation
echo   - Settings
echo   - Exit
echo.

start "" powershell.exe -NoExit -Command "Set-Location '%~dp0'; dotnet run -c Release --framework net10.0-windows -- listen"
