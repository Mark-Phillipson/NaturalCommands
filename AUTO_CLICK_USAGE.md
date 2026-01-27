# Auto-Click Feature Usage Guide

## Overview
The auto-click feature automatically performs a left-click when your mouse stays still for a configurable duration (default 500ms). This is useful for hands-free gaming and other scenarios where you want to click without manual intervention.

## How to Use

### Starting the Application in Listen Mode

**IMPORTANT**: The auto-click feature requires the application to run in **listen mode** to work properly.

#### Option 1: Run with `/listen` argument
```powershell
.\NaturalCommands.exe /listen
```

#### Option 2: Create a shortcut for listen mode
1. Right-click on `NaturalCommands.exe` → Create Shortcut
2. Right-click the shortcut → Properties
3. In the "Target" field, add `/listen` at the end:
   ```
   "C:\Path\To\NaturalCommands.exe" /listen
   ```
4. Click OK
5. Double-click the shortcut to start in listen mode

When running in listen mode, you'll see:
- **System tray icon** in the notification area (bottom-right of screen)
- Right-click the icon to access:
  - **Open Voice Dictation** - Voice command interface
  - **Settings...** - Configure all settings including auto-click
  - **Exit** - Close the application

### Enabling Auto-Click

Once the application is running in listen mode, you can enable auto-click in two ways:

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
- **"stop auto click"** or **"disable auto click"**

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

## Troubleshooting

### Auto-click not working:
- ✓ Ensure you're running in **listen mode** (`/listen`), not direct command mode
- ✓ Check that the tray icon is visible in the system notification area
- ✓ Verify auto-click is enabled (use voice command "auto click")
- ✓ Test with the overlay visible to see the countdown

### Tray icon not showing:
- ✓ Run with `/listen` argument
- ✓ Check if hidden in the overflow area (click ^ icon in tray)
- ✓ Restart the application
- ✓ Check Windows notification area settings

### Clicks happening too fast/slow:
- Open **Settings** → **Auto-Click** tab
- Adjust the delay (100-2000ms)
- Save and test again

## Advanced Usage

### Starting Listen Mode on Windows Startup

1. Press **Win+R**, type `shell:startup`, press Enter
2. Create a shortcut to `NaturalCommands.exe` with `/listen` argument
3. The application will start automatically when Windows boots

### Using with Games

1. Start NaturalCommands in listen mode
2. Launch your game
3. Use voice command to enable auto-click: **"auto click"**
4. The auto-click will work across all applications
5. To disable during gameplay: **"stop auto click"**

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
