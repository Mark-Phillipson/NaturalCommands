## Talon Voice Integration

NaturalCommands can be triggered using Talon Voice by saying **"natural"** (or **"nat"**) followed by the command you want to execute.

### Main Commands

```talon
^(natural | nat) <user.text>$:
    user.run_application_csharp_natural(text)
^(natural | nat) dictate$:
    speech.disable()
    user.run_application_csharp_natural("dictate")
^show letters$:
    user.run_application_csharp_natural("show letters")
^show taskbar$:
    user.run_application_csharp_natural("show taskbar")
^show desktop$:
    user.run_application_csharp_natural("show desktop")
^(ID) <user.text>$:
    user.run_application_csharp_natural("identify " + text)
```

**Voice command examples:**
- `natural what can I say` - display available commands
- `natural dictate` - open voice dictation UI
- `show letters` - overlay letter labels on clickable elements
- `show taskbar` - focus and label taskbar icons
- `show desktop` - focus and label desktop icons
- `identify button` - find and highlight target elements

### Semantic Command Matching

NaturalCommands supports **semantic matching** for Talon voice commands. This means if you say something that's **semantically similar** but not an exact textual match to a registered command, it will still be evaluated and executed based on meaning.

**How it works:**
1. When you issue a voice command via Talon, NaturalCommands first attempts an exact match against known commands
2. If no exact match is found, the fuzzy matcher scores the input against available commands by similarity
3. If a good semantic match is found (e.g., "close tab" matches "close the current tab"), it will be executed
4. If still no suitable match, the AI interpreter is invoked as a fallback to infer the intended action

**Example:**
- Registered command: `close tab`
- You say: `natural close this tab`
- Result: Matches semantically and executes the close tab action

### Using `mimic` with Talon Commands (Workaround)

As a workaround in custom Talon my stuff repository, the `mimic` command is used to automatically invoke voice commands when a command phrase has been recognized as having Talon command semantics (even though `mimic` is not formally supported by Talon). This is the only available workaround for executing a NaturalCommands action based on voice input that was processed with the same semantic meaning.

The **talon_stuff repository** automatically scans a text file that lists command phrases, and generates `mimic` calls to execute them through NaturalCommands. This scanning mechanism allows commands to be executed based on their semantic recognition rather than exact Talon command matching.

### Python Code

```python
def run_application_csharp_natural(naturalCommand: str):
	"""runs the natural command with the given text"""
	import os
	commandline = r'C:\Users\MPhil\source\repos\NaturalCommands\bin\Release\net10.0-windows\NaturalCommands.exe'
	args1 = ' ' + r'/natural' + ' '
	args2 = '' + r'/' + naturalCommand + ''
	arguments = [args1, args2]
	cwd = os.path.dirname(commandline)
	print(commandline)
	print(naturalCommand)
	ui.launch(path=commandline, args=arguments, cwd=cwd)
```

### Mouse Movement Commands

NaturalCommands includes Talon voice commands for continuous mouse movement.

**Features:**
- Directional movement: left, right, up, down, and diagonals
- Speed adjustment: "mouse faster" or "mouse slower"
- Stop commands: "mouse stop", "stop click", "stop right click"

**Mouse Commands:**

```talon
# Basic directional movement
^mouse left$:
    user.run_application_csharp_natural("mouse left")
^mouse right$:
    user.run_application_csharp_natural("mouse right")
^mouse up$:
    user.run_application_csharp_natural("mouse up")
^mouse down$:
    user.run_application_csharp_natural("mouse down")

# Diagonal movement
^mouse up left$:
    user.run_application_csharp_natural("mouse up left")
^mouse up right$:
    user.run_application_csharp_natural("mouse up right")
^mouse down left$:
    user.run_application_csharp_natural("mouse down left")
^mouse down right$:
    user.run_application_csharp_natural("mouse down right")

# Stop commands
^mouse stop$:
    user.run_application_csharp_natural("mouse stop")
^stop mouse$:
    user.run_application_csharp_natural("stop mouse")
^stop click$:
    user.run_application_csharp_natural("stop click")
^stop right click$:
    user.run_application_csharp_natural("stop right click")

# Speed adjustment
^mouse faster$:
    user.run_application_csharp_natural("mouse faster")
^mouse slower$:
    user.run_application_csharp_natural("mouse slower")
^faster$:
    user.run_application_csharp_natural("faster")
^slower$:
    user.run_application_csharp_natural("slower")
```

**Voice command examples:**
- `mouse left` - start moving the mouse left
- `mouse up right` - move diagonally up and right
- `mouse faster` - increase movement speed
- `mouse stop` - stop all mouse movement

# NaturalCommands

Execute Windows and Visual Studio commands using natural language. If a command is not recognized, the AI will infer the most likely intended command, so you are not limited to a single phrase—commands are interpreted by meaning.

## Overview

NaturalCommands is a lightweight Windows application that maps natural language phrases to system and Visual Studio actions. It includes helpers for sending keys, managing windows, running processes, and integrating with AI-based interpreters.

## Features

 - Map natural language to keyboard and window actions
 - Voice dictation helpers and multi-action support
 - Visual Studio command helpers and shortcuts
 - **Show letters** feature for voice-based UI element navigation
 - **Flexible command interpretation:** If a command is not recognized, the AI will decide or figure out the likely command based on the meaning, so commands are not limited to a single phrase.

## Voice Command: "What can I say?"

You can say **"what can I say"** at any time to display a list of available commands.

## Prerequisites

- Windows 10 or later
- [.NET SDK 10.0+](https://dotnet.microsoft.com/en-us/download)

## Build & Run

```bash
# Restore, build and run
dotnet build NaturalCommands.csproj  
dotnet build NaturalCommands.csproj  -c Release
```

## Listen mode (resident hotkey)

Run the app in resident mode to open the voice dictation UI from anywhere via a global hotkey.

- Start resident mode:
	- `dotnet run --framework net10.0-windows -- listen`
	- or run the built exe: `NaturalCommands.exe listen`
- Hotkey: **Win+Ctrl+H**
- Fallback: use the system tray menu item **Open Voice Dictation (Win+Ctrl+H)**

When voice typing stops, the **Send Command** button is focused and the form’s **Enter** key submits (AcceptButton), so you can press Enter to send the captured text.

## Show Letters Feature 🎯

The "show letters" feature enables voice-based navigation of UI elements by overlaying letter labels on clickable items.

### Usage

1. Say **"show letters"** or **"natural show letters"**
2. A transparent overlay will appear showing two-letter labels (a, b, c, ..., aa, ab, ...) on all clickable elements
3. Type the letters corresponding to the element you want to click using the Talon alphabet
4. The element will be clicked automatically
5. Press **ESC** to cancel without clicking

### Examples

```talon
# Show letter labels on clickable elements
^show letters$:
    user.run_application_csharp_natural("show letters")

# Then type letters like "a", "b", "ab", etc. to click the corresponding element
# Voice command: say "show letters"
```

### Show Desktop

Talon command:
```talon
^show desktop$:
    user.run_application_csharp_natural("show desktop")
```

Usage:
1. Say **"show desktop"**
2. The desktop will be focused (icons exposed) and letter labels will be applied to desktop icons so you can activate them by typing the letters; press **ESC** to cancel.

### Supported Elements

The feature works with:
- Buttons
- Links
- Menu items
- Checkboxes
- Radio buttons
- Tab items
- Combo boxes

### Works With

- **Web browsers** (Edge, Chrome, Firefox)
- **Desktop applications** (Windows native apps)
- Any application that supports Windows UI Automation API

## Visual Target Click (identify / choose)

NaturalCommands supports a voice-driven visual targeting flow for commands like identifying and clicking specific on-screen targets.

### Talon Commands

```talon
^(ID) <user.text>$:
    user.run_application_csharp_natural("identify " + text)
```

Voice command usage:
- `identify <target>`: Find matching targets on screen.
- `show candidates`: Re-show the numbered overlay for the latest identify session.
- `choose <number>`: Click a numbered candidate from the current session.

### Behavior

1. Local UI Automation is attempted first.
2. If local confidence is not high enough and cloud vision is enabled, a screenshot is sent to OpenAI vision.
3. If one strong candidate is found, NaturalCommands auto-clicks it.
4. If multiple candidates are found, a numbered overlay is shown and you can say `choose 1`, `choose 2`, etc.

### Settings

Open **Settings → AI Integration → Visual Targeting**:

- **Auto-click confidence threshold** (`0.50` to `1.00`)
- **Max cloud visual calls per day** (`0` disables cloud calls)
- **Model tier** (`off`, `fast`, `balanced`)
- **Fallback mode** (`uia-then-ocr`, `uia-only`, `ocr-only`)

Cloud budget counters are tracked daily (UTC) in `settings.json` and reset automatically each day.

## Auto-Click (countdown overlay)

Auto-Click provides a configurable delayed-click feature that displays a small countdown overlay centered on the cursor during the delay. While the countdown is active the overlay temporarily hides the system cursor and shows a progress arc and a small hotspot dot so the user can clearly see where the click will occur.

**Key points:**
- The auto-click delay is configurable via **Settings → Auto-Click** (milliseconds).
- The **Show countdown overlay** setting toggles whether the overlay is shown during the delay.
- When shown, the overlay replaces (hides) the system cursor during the countdown and restores it when the countdown finishes or is cancelled.
- There is a **Stop Auto-Click** tray menu option to cancel the countdown, and the auto-click will also automatically stop if the mouse is moved to the taskbar/system tray to prevent accidental clicks.

**Talon Command:**
```talon
^auto click$:
    user.run_application_csharp_natural("auto click")
```

**Usage:**
1. Open **Settings → Auto-Click** to enable/set the delay and toggle the overlay.
2. Say **"auto click"** to trigger a delayed click; the overlay will appear centered on the pointer and display progress.
3. Cancel with the **Stop Auto-Click** tray menu item or by moving the pointer to the taskbar.

## Contributing

Contributions are welcome — please open issues and pull requests on GitHub.

## License

This project is licensed under the MIT License — see the `LICENSE` file for details.
