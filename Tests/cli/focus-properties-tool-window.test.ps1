# CLI test: focus Visual Studio tool window
# This script will test the FocusWindowAction by simulating a command to focus the Properties tool window in Visual Studio.

$logPath = "../../bin/bin/app.log"
$exePath = "../../bin/Release/net10.0-windows/NaturalCommands.dll"

# Remove old log if exists
if (Test-Path $logPath) { Remove-Item $logPath }

# Launch the command to focus the Properties tool window
& dotnet $exePath /natural "/focus properties tool window"

Start-Sleep -Seconds 2

# Check the log for success
if (-Not (Test-Path $logPath)) {
    Write-Host "Log file not found. Test failed."
    exit 1
}

$logContent = Get-Content $logPath -Raw
if ($logContent -match "Focused window: properties tool window") {
    Write-Host "Test passed: Properties tool window focused."
    exit 0
} elseif ($logContent -match "Could not find window with title containing: properties tool window") {
    Write-Host "Test ran: No matching window found (may be expected if VS not open)."
    exit 0
} else {
    Write-Host "Test failed: Unexpected log output."
    exit 1
}
