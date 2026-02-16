# Quick Clicks Show/Hide Testing Guide

## ✨ Updated Behavior (February 2026)

**Quick Click overlays are now HIDDEN BY DEFAULT.** They will only appear when you explicitly request them via the "show quick clicks" command. This provides better control and prevents unwanted overlays from appearing automatically.

- ✅ Overlays start hidden
- ✅ Say "show quick clicks" to display them
- ✅ Say "hide quick clicks" to hide them  
- ✅ Optional: Enable `AutoShowOnFocus` in settings if you want auto-show behavior

### Important Note for Existing Users

If you have an existing `settings.json` file from before this change, it may still have `"AutoShowOnFocus": true`. To get the new default behavior:

**Option 1 - Quick Fix:** Delete your `settings.json` file to regenerate with new defaults
```powershell
Remove-Item "bin\Release\net10.0-windows\settings.json" -Force
```

**Option 2 - Manual Edit:** Open `settings.json` and change:
```json
"QuickClicks": {
  "AutoShowOnFocus": false
}
```

## Problem Diagnosis

Based on the log file, `hide quick clicks` is working correctly:
- Manual hide flag file is created
- Overlay is hidden

The issue with `show quick clicks` is likely one of:
1. **No matching profile exists** for the current application
2. **Monitor resolution mismatch** (profile exists for different resolution)
3. **Manual hide flag not being cleared** (but code shows it should be)

## Test Scenarios

### Test 0: Verify New Default Behavior (Recommended First Test)

This test confirms overlays are hidden by default:

```powershell
# 1. Create a test profile for Notepad
Set-Content "quick_clicks.json" -Value @'
[
  {
    "ProcessName": "notepad",
    "WindowTitlePattern": null,
    "MonitorWidth": null,
    "MonitorHeight": null,
    "Regions": [
      {
        "Name": "Test",
        "X": 200,
        "Y": 200,
        "Width": 100,
        "Height": 50,
        "ClickType": "Left"
      }
    ]
  }
]
'@

# 2. Open Notepad
Start-Process notepad.exe
Start-Sleep -Seconds 2

# 3. Verify overlay does NOT auto-show (new default behavior)
# You should NOT see any overlay at this point

# 4. Explicitly show the overlay
.\bin\Release\net10.0-windows\NaturalCommands.exe /natural "show quick clicks"
# Now you SHOULD see a blue overlay box labeled "Test" at position 200,200

# 5. Hide it
.\bin\Release\net10.0-windows\NaturalCommands.exe /natural "hide quick clicks"
# Overlay should disappear

# 6. Show it again
.\bin\Release\net10.0-windows\NaturalCommands.exe /natural "show quick clicks"
# Overlay should reappear
```

**Expected Results:**
- ✅ Step 2-3: No overlay appears automatically (this is the new behavior)
- ✅ Step 4: Overlay appears when explicitly shown
- ✅ Step 5: Overlay disappears when hidden
- ✅ Step 6: Overlay reappears when shown again

### Test 1: Check Current State

```powershell
# Check if manual hide flag exists
Get-Item "bin\Release\net10.0-windows\.quick_clicks_manually_hidden" -ErrorAction SilentlyContinue

# Check if quick_clicks.json exists and what's in it
Get-Content "quick_clicks.json" -ErrorAction SilentlyContinue | ConvertFrom-Json
```

### Test 2: Create a Test Profile

Create or update `quick_clicks.json` with a test profile for your current application:

```json
[
  {
    "ProcessName": "devenv",
    "WindowTitlePattern": null,
    "MonitorWidth": null,
    "MonitorHeight": null,
    "Regions": [
      {
        "Name": "TestRegion",
        "X": 100,
        "Y": 100,
        "Width": 150,
        "Height": 50,
        "ClickType": "Left",
        "VoiceCommand": "test region"
      }
    ]
  }
]
```

**Note:** Replace `"devenv"` with the process name of the app you're testing with (e.g., `"notepad"`, `"chrome"`, `"code"`).

### Test 3: Manual Flag Management Test

Run these commands in sequence from the command line:

```powershell
# 1. Hide quick clicks
.\bin\Release\net10.0-windows\NaturalCommands.exe /natural "hide quick clicks"

# Check flag was created
Test-Path "bin\Release\net10.0-windows\.quick_clicks_manually_hidden"
# Should return: True

# 2. Show quick clicks
.\bin\Release\net10.0-windows\NaturalCommands.exe /natural "show quick clicks"

# Check flag was deleted (cleared)
Test-Path "bin\Release\net10.0-windows\.quick_clicks_manually_hidden"
# Should return: False (if show worked correctly)
```

### Test 4: Verbose Logging Test

Add verbose logging to see what's happening in ShowQuickClicksAction:

```powershell
# Run with the app focused (e.g., Notepad)
Start-Process notepad.exe
Start-Sleep -Seconds 2

# Try to show quick clicks for notepad
.\bin\Release\net10.0-windows\NaturalCommands.exe /natural "show quick clicks"

# Check the log
Get-Content "bin\bin\app.log" -Tail 50
```

Look for:
- "Quick Clicks shown." - Success
- "Quick Clicks mismatch warning shown." - Profile exists but wrong resolution
- "No quick-click profiles for current app." - No profile found

## Expected Behavior (Updated - Default is Hidden)

**As of the latest update, overlays are HIDDEN BY DEFAULT** and only appear when explicitly requested via "show quick clicks" command.

### Working correctly:
1. **Default state**: Overlays do NOT auto-show when apps gain focus (AutoShowOnFocus = false by default)
2. **Show**: Clears manual hide flag, shows overlay if profile exists for current app
3. **Hide**: Sets manual hide flag, overlay disappears
4. **Auto-show disabled**: ForegroundMonitor will NOT automatically show overlays unless AutoShowOnFocus is enabled in settings

### Common Issues:

#### Issue 1: No Profile Exists
**Symptom:** "No quick-click profiles for current app." in execution result  
**Fix:** Create a profile for the app (see Test 2 above)

#### Issue 2: Resolution Mismatch
**Symptom:** "Quick Clicks mismatch warning shown."  
**Fix:** Either:
- Set `MonitorWidth` and `MonitorHeight` to `null` in the profile (works on any monitor)
- Create a profile with the exact monitor resolution

#### Issue 3: Manual Hide Flag Stuck
**Symptom:** Show command runs but overlay doesn't appear, flag file still exists  
**Fix:** Manually delete the flag file:
```powershell
Remove-Item "bin\Release\net10.0-windows\.quick_clicks_manually_hidden" -Force
```

## Testing Workflow

1. **Setup**: Create a test profile for a simple app like Notepad
2. **Focus the app**: Make sure Notepad is the foreground window
3. **Test hide**: Run hide command, verify overlay disappears
4. **Test show**: Run show command, verify overlay reappears
5. **Check logs**: Examine app.log for detailed execution flow

## Quick Diagnosis Commands

```powershell
# All-in-one diagnostic script
Write-Host "=== Quick Clicks Diagnostic ===" -ForegroundColor Cyan

# Check manual hide flag
$flagExists = Test-Path "bin\Release\net10.0-windows\.quick_clicks_manually_hidden"
Write-Host "Manual hide flag exists: $flagExists" -ForegroundColor $(if($flagExists){'Yellow'}else{'Green'})

# Check profiles
if (Test-Path "quick_clicks.json") {
    $profiles = Get-Content "quick_clicks.json" | ConvertFrom-Json
    Write-Host "Number of profiles: $($profiles.Count)" -ForegroundColor Green
    foreach ($p in $profiles) {
        Write-Host "  - Process: $($p.ProcessName), Regions: $($p.Regions.Count), Monitor: $($p.MonitorWidth)x$($p.MonitorHeight)" -ForegroundColor Gray
    }
} else {
    Write-Host "quick_clicks.json NOT FOUND" -ForegroundColor Red
}

# Check recent log entries
Write-Host "`nRecent log entries:" -ForegroundColor Cyan
if (Test-Path "bin\bin\app.log") {
    Get-Content "bin\bin\app.log" -Tail 10 | Where-Object { $_ -match "QuickClick" }
}
```

## Recommended Fix

If show/hide isn't working, the most likely issue is no matching profile. Create a universal test profile:

```json
[
  {
    "ProcessName": "notepad",
    "WindowTitlePattern": null,
    "MonitorWidth": null,
    "MonitorHeight": null,
    "Regions": [
      {
        "Name": "Test",
        "X": 200,
        "Y": 200,
        "Width": 100,
        "Height": 50,
        "ClickType": "Left"
      }
    ]
  }
]
```

Then:
1. Open Notepad
2. Run: `.\bin\Release\net10.0-windows\NaturalCommands.exe /natural "show quick clicks"`
3. You should see a blue overlay box labeled "Test" at position 200,200
4. Run: `.\bin\Release\net10.0-windows\NaturalCommands.exe /natural "hide quick clicks"`
5. Overlay should disappear
6. Run show again - should reappear
