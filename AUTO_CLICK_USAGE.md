# Auto-Click Feature Usage Guide

## Overview
The auto-click feature automatically performs a left-click when your mouse stays still for a configurable duration (default 500ms). This is useful for hands-free gaming and other scenarios where you want to click without manual intervention.

## How to Use

### Starting the Application (resident tray/menu)

**IMPORTANT**: The auto-click feature requires the application to be running in resident/tray mode (the app shows a system tray icon and registers the hotkey). By default the app creates the tray/menu at startup when `Settings -> Notifications -> ShowTrayOnStartup` is enabled.

#### Option 1: Run with no arguments (default resident mode)
```powershell
.\NaturalCommands.exe
```

#### Option 2: Create a shortcut that starts the app
1. Right-click on `NaturalCommands.exe` → Create Shortcut
2. Right-click the shortcut → Properties
3. Ensure the "Target" field points to the exe (no extra args required):
  ```
  "C:\Path\To\NaturalCommands.exe"
  ```
4. Click OK
5. Double-click the shortcut to start the app and show the tray icon

When running in resident/tray mode, you'll see:
- **System tray icon** in the notification area (bottom-right of screen)
- Right-click the icon to access:
  - **Open Voice Dictation** - Voice command interface
  - **Settings...** - Configure all settings including auto-click
  - **Exit** - Close the application

### Enabling Auto-Click

Once the application is running in resident/tray mode, you can enable auto-click in two ways:

#### Method 1: Voice Command
1. Press **Win+Ctrl+H** to open voice dictation
2. Say or type: **"auto click"**
3. Press Submit

#### Method 2: Direct Command (from another instance)
```powershell
.\NaturalCommands.exe /natural "auto click"
```

### Disabling Auto-Click

#### Voice Command:
- **"stop auto click"**, **"disable auto click"**, or **"auto click off"**

#### Adjusting Auto-Click Speed (voice)
- **"auto click faster"** or **"auto click speed up"** — decrease the delay between idle and click (makes clicks occur sooner)
- **"auto click slower"** or **"auto click slow down"** — increase the delay between idle and click (makes clicks occur later)

Notes: Speed adjustments are applied in steps (see settings for allowed range).
### Configuring Auto-Click

1. Right-click the **system tray icon**
2. Select **"Settings..."**
3. Go to the **"Auto-Click"** tab
4. Adjust settings:
   - **Idle delay before click (ms)**: 100-2000ms (default: 500ms)
   - **Show countdown overlay**: Enable/disable the blue circle indicator
5. Click **"Save"** or **"Apply"**

## Visual Feedback

When auto-click is active and the mouse is idle:
- A **blue circular progress indicator** appears near your cursor
- Shows the **percentage** countdown (0% to 100%)
- Displays **remaining milliseconds**
- The overlay disappears when you move the mouse

### Troubleshooting

### Auto-click not working:
- ✓ Ensure the app is running in resident/tray mode (no extra args required)
- ✓ Check that the tray icon is visible in the system notification area
- ✓ Verify auto-click is enabled (use voice command "auto click")
- ✓ Test with the overlay visible to see the countdown

### Tray icon not showing:
- ✓ Run the app (no extra args required)
- ✓ Check if hidden in the overflow area (click ^ icon in tray)
- ✓ Restart the application
- ✓ Check Windows notification area settings

### Clicks happening too fast/slow:
- Open **Settings** → **Auto-Click** tab
- Adjust the delay (100-2000ms)
- Save and test again

## Advanced Usage

### Starting Resident Mode on Windows Startup

1. Press **Win+R**, type `shell:startup`, press Enter
2. Create a shortcut to `NaturalCommands.exe` (no extra args required)
3. The application will start automatically when Windows boots

### Using with Games

1. Start NaturalCommands (resident/tray mode)
2. Launch your game
3. Use voice command to enable auto-click: **"auto click"**
4. The auto-click will work across all applications
5. To disable during gameplay: **"stop auto click"**

## Visual Target Click Usage

This feature is useful when you want to click a specific visible target by description.

### Voice commands

- **"identify ten of spades"**
- **"show candidates"**
- **"choose 2"**

### What happens

1. NaturalCommands tries local UI Automation first.
2. If needed and cloud calls are enabled, it sends a compressed all-monitor screenshot to OpenAI vision.
3. A single high-confidence target is auto-clicked.
4. Multiple matches show a numbered overlay for `choose <number>`.

### Cost and latency controls

Use **Settings → AI Integration → Visual Targeting** to set:

- Max cloud calls/day
- Model tier (`off`, `fast`, `balanced`)
- Confidence threshold for auto-click
- Fallback mode (`uia-then-ocr`, `uia-only`, `ocr-only`)

## Settings File

Settings are stored in:
```
<AppDirectory>\settings.json
```

You can edit this file directly (while the app is not running) or use the Settings form in the tray menu.

## Default Settings

```json
{
  "AutoClick": {
    "DelayMs": 500,
    "Enabled": false,
    "ShowOverlay": true
  }
}
```

Default notification settings (new):

```json
{
  "Notifications": {
    "DisplayDurationMs": 5000,
    "Enabled": true,
    "TickerEnabled": false,
    "MaxTooltipLength": 60
  }
}
```
