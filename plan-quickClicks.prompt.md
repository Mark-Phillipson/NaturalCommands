## Plan: Quick Clicks — Named Clickable Region Overlays

**TL;DR**: Add a "Quick Clicks" system that lets users define named clickable regions per app (and optionally per window title). Regions are stored in `quick_clicks.json`, auto-shown when their associated app gains focus, hideable/showable via voice, and editable in-place. Saying a region's name clicks its center. The system auto-generates `.talon` files for voice integration. A foreground-window monitor (`SetWinEventHook`) in `ListenModeApplicationContext` drives auto-show/hide. The overlay reuses the transparent fullscreen WinForms pattern from `UIElementOverlayForm`. An in-place edit mode allows dragging regions, resizing via buttons, and creating new regions from the current mouse position.

**Steps**

### 1. Data Model — `Models/QuickClickRegion.cs` (new file)
- Define a `QuickClickProfile` class: `ProcessName` (string), `WindowTitlePattern` (string?, nullable for "any title"), `Regions` (list of `QuickClickRegion`).
- Define a `QuickClickRegion` class: `Name` (string), `X` (int), `Y` (int), `Width` (int, default 100), `Height` (int), `VoiceCommand` (string? — defaults to `Name` lowered).
- Coordinates are **relative to the app window's client area** (not absolute screen coords) so they survive window moves/resizes. Store as **pixel offsets from the window top-left** since most apps are used at consistent sizes, and percentages would be hard to reason about visually. We'll document that users should re-adjust after major window resizes.

### 2. Storage — `quick_clicks.json` + `Helpers/QuickClickLoader.cs` (new files)
- JSON file at the app directory, array of `QuickClickProfile` objects. Example:
  ```json
  [
    {
      "ProcessName": "chrome",
      "WindowTitlePattern": "YouTube",
      "Regions": [
        { "Name": "Like", "X": 450, "Y": 620, "Width": 100, "Height": 40 }
      ]
    }
  ]
  ```
- `QuickClickLoader` — static class modeled after `Helpers/MultiActionLoader.cs`. Methods: `Load()`, `Save(profiles)`, `GetProfilesForApp(processName, windowTitle)` which returns matching profiles (exact process match, optional title substring/regex match). Loaded at startup in `NaturalLanguageInterpreter` static constructor alongside `MultiActionLoader.Load()`.

### 3. Action Model — add to `ActionModels.cs`
- Add `public record ClickQuickClickAction(string RegionName) : ActionBase;`
- Add `public record ShowQuickClicksAction() : ActionBase;`
- Add `public record HideQuickClicksAction() : ActionBase;`
- Add `public record EditQuickClicksAction() : ActionBase;`
- Add `public record CreateQuickClickAction() : ActionBase;` — for creating from current mouse position.

### 4. Foreground Window Monitor — `Helpers/ForegroundMonitor.cs` (new file)
- Use `SetWinEventHook(EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND, IntPtr.Zero, callback, 0, 0, WINEVENT_OUTOFCONTEXT)` to detect focus changes.
- Add required P/Invoke signatures: `SetWinEventHook`, `UnhookWinEvent`, and public `GetForegroundWindow` / `GetWindowText` to `Helpers/Win32ApiHelper.cs`.
- On each focus change: get process name via `CurrentApplicationHelper.GetCurrentProcessName()`, get window title via new `CurrentApplicationHelper.GetCurrentWindowTitle()`, check `QuickClickLoader.GetProfilesForApp(...)`. If profiles exist → show `QuickClickOverlayForm`. If no profiles → hide it.
- Debounce rapid focus changes (100-200ms) using `DebounceGate` from `Helpers/DebounceGate.cs` or a simple timer.
- **Critical**: The overlay itself must NOT steal focus (use `WS_EX_NOACTIVATE` + `WS_EX_TOOLWINDOW` extended styles), otherwise it will trigger its own focus-change event in an infinite loop.
- Initialize in `ListenModeApplicationContext` constructor after hotkey registration (~line 87). Unhook in `ExitThreadCore()`.

### 5. Extend `CurrentApplicationHelper` — `CurrentApplicationHelper.cs`
- Add `public static string? GetCurrentWindowTitle()` — `GetForegroundWindow()` → `GetWindowText()` → return the title string. Reuse the `GetWindowText` P/Invoke already in `WindowFocusHelper` but make it accessible here.

### 6. Overlay Form — `QuickClickOverlayForm.cs` (new file)
- **Display mode** (non-interactive overlay):
  - Transparent fullscreen WinForms, same pattern as `UIElementOverlayForm.cs` (`TransparencyKey = Color.Lime`, borderless, topmost).
  - **Must be click-through** in display mode using `WS_EX_TRANSPARENT | WS_EX_LAYERED | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW` (unlike `UIElementOverlayForm` which captures input). This ensures the user can interact with the underlying app normally.
  - Draw each region as a semi-transparent colored rectangle with the region name as a label (GDI+, anti-aliased, similar styling to `UIElementOverlayForm`'s label rendering).
  - Convert window-relative coordinates to screen coordinates using the target window's position (`GetWindowRect`).
  - Singleton pattern (like `AutoClickOverlayForm`) with static `Show(profiles, windowHandle)`, `Hide()`, `IsVisible` methods. Thread-safe via `SynchronizationContext`.
  - Refresh/reposition when the target app window moves (poll via timer at ~200ms or hook `EVENT_OBJECT_LOCATIONCHANGE`).

- **Edit mode** (interactive):
  - Toggled by voice command ("edit quick clicks") or a small gear button drawn on the overlay.
  - Switch from click-through to interactive: remove `WS_EX_TRANSPARENT` flag, set `WS_EX_NOACTIVATE` to keep focus behavior correct.
  - Each region becomes **draggable** — mouse down on a region starts drag, mouse move repositions it, mouse up commits.
  - **Selection**: clicking a region selects it (highlighted border). Selected region shows **resize buttons**: `[Wider]` `[Narrower]` `[Taller]` `[Shorter]` rendered as small button-like rectangles near the selected region. Each adjusts width/height by ±20px per click.
  - **New region from mouse**: a `[+ New]` button on the overlay. Clicking it enters "placement mode" — next click on the overlay captures mouse position, prompts for a name (small textbox popup or `InputBox`-style dialog), creates a 100×100 region at that point.
  - **Rectangle area mode**: a `[+ Rectangle]` button. First click sets corner 1, second click sets corner 2, then prompt for name. This defines width/height from the two corners.
  - **Delete**: a `[Delete]` button visible when a region is selected.
  - **Done/Save**: a `[Save & Exit]` button that saves to `quick_clicks.json` via `QuickClickLoader.Save()` and switches back to display mode.
  - All edit UI rendered with GDI+ (no child WinForms controls, to keep the transparent overlay clean). Hit detection via manual coordinate checks in mouse event handlers.

### 7. Command Routing — `NaturalLanguageInterpreter.cs`
- In `InterpretAsync`, add a block (before the AI fallback, after emoji/general commands):
  - "show quick clicks" → `ShowQuickClicksAction`
  - "hide quick clicks" → `HideQuickClicksAction`
  - "edit quick clicks" → `EditQuickClicksAction`
  - "create quick click" → `CreateQuickClickAction`
  - **Dynamic region name matching**: load current profiles for the focused app, check if `text` matches any region's `Name` (case-insensitive, normalized) → `ClickQuickClickAction(regionName)`.
- In `ExecuteActionAsync` (same file, ~line 1161), add handlers:
  - `ClickQuickClickAction` → look up region in current app's profile → calculate screen coordinates from window position + region offset → `mouse_event` at center point (reuse `SetCursorPos` + `mouse_event(LEFTDOWN/LEFTUP)` pattern from `AutoClickManager`).
  - `ShowQuickClicksAction` → `QuickClickOverlayForm.Show(...)`.
  - `HideQuickClicksAction` → `QuickClickOverlayForm.Hide()`.
  - `EditQuickClicksAction` → `QuickClickOverlayForm.EnterEditMode()`.
  - `CreateQuickClickAction` → get current cursor position, prompt for name via a small `InputBox` form, create a 100×100 region, add to current profile, save, refresh overlay.

### 8. Create Quick Click from Mouse Position — `CreateQuickClickForm.cs` (new, small)
- Minimal input dialog: a textbox for the region name + OK/Cancel buttons.
- Shown centered on screen when "create quick click" command is executed.
- After OK: creates a `QuickClickRegion` at the saved cursor position (captured *before* showing the dialog), adds to the matching profile (or creates a new profile for the current app), saves, and refreshes the overlay.

### 9. Talon File Generation — `Helpers/QuickClickTalonGenerator.cs` (new file)
- After any save to `quick_clicks.json`, auto-generate `.talon` files.
- One file per unique `ProcessName`: e.g., `quick_clicks_chrome.talon`, `quick_clicks_devenv.talon`.
- File placed in a configurable Talon user directory (add `TalonUserDirectory` setting to `Models/AppSettings.cs`) or alongside the app in a `talon_output/` folder.
- Format per region:
  ```
  ^click like$:
      user.run_application_csharp_natural("like")
  ```
- Include a header comment: `# Auto-generated by NaturalCommands Quick Clicks — do not edit manually`.
- Regenerate fully on each save (overwrite).

### 10. Settings Integration — `Models/AppSettings.cs` + `SettingsForm.cs`
- Add a `QuickClickSettings` section to `AppSettings`:
  - `Enabled` (bool, default true)
  - `AutoShowOnFocus` (bool, default true)
  - `OverlayOpacity` (float, 0.0-1.0, default 0.5)
  - `DefaultRegionSize` (int, default 100)
  - `TalonOutputDirectory` (string?, nullable — if null, output to app directory)
- Add a "Quick Clicks" tab to `SettingsForm` with controls for these settings.

### 11. Tray Menu Integration — `ListenModeApplicationContext.cs`
- Add a "Quick Clicks" submenu to the tray context menu with items:
  - "Show Quick Clicks" → triggers `ShowQuickClicksAction`
  - "Edit Quick Clicks" → triggers `EditQuickClicksAction`
  - "Create Quick Click at Mouse" → triggers `CreateQuickClickAction`
  - "Open quick_clicks.json" → opens the JSON file in default editor

### 12. New Files Summary

| File | Purpose |
|---|---|
| `Models/QuickClickRegion.cs` | Data model classes |
| `Helpers/QuickClickLoader.cs` | JSON load/save |
| `Helpers/ForegroundMonitor.cs` | `SetWinEventHook` focus tracking |
| `QuickClickOverlayForm.cs` | Display + edit overlay |
| `CreateQuickClickForm.cs` | Name input dialog |
| `Helpers/QuickClickTalonGenerator.cs` | `.talon` file generation |

### 13. Modified Files Summary

| File | Changes |
|---|---|
| `ActionModels.cs` | Add 5 new action records |
| `NaturalLanguageInterpreter.cs` | Add command matching + execution handlers |
| `CurrentApplicationHelper.cs` | Add `GetCurrentWindowTitle()` |
| `Helpers/Win32ApiHelper.cs` | Add `SetWinEventHook`, `UnhookWinEvent`, `GetForegroundWindow`, `GetWindowText` P/Invokes |
| `ListenModeApplicationContext.cs` | Initialize `ForegroundMonitor`, add tray menu items |
| `Models/AppSettings.cs` | Add `QuickClickSettings` section |
| `SettingsForm.cs` | Add Quick Clicks settings tab |

### Verification
1. Build with `dotnet build NaturalCommands.csproj` — should compile clean.
2. Run in listen mode (`/listen`), open Chrome → verify overlay auto-appears if a chrome profile exists in `quick_clicks.json`.
3. Say "edit quick clicks" → verify edit mode activates with draggable regions and resize buttons.
4. Click `[+ New]` → click on overlay → enter name → verify region created and saved.
5. Say "hide quick clicks" → verify overlay hides. Say "show quick clicks" → verify it returns.
6. Say a region name (e.g., "like") → verify cursor moves to region center and clicks.
7. Switch to a different app → verify overlay hides. Switch back → verify it re-appears.
8. Check that `.talon` files are generated in the output directory with correct command format.
9. Test with window title filtering — define a profile for `chrome` + `YouTube` title, verify overlay only shows on YouTube tabs.

### Design Decisions
- **Pixel offsets from window top-left** over percentage-based coordinates — simpler to reason about, matches how users position things visually. Trade-off: regions won't auto-adapt to window resize, but a re-edit is easy.
- **`SetWinEventHook`** over timer polling — more responsive, lower CPU usage, no polling overhead.
- **GDI+ rendered buttons** in edit mode over WinForms child controls — keeps the transparent overlay architecture clean and avoids focus/z-order issues with child controls on transparent forms.
- **Click-through in display mode** (`WS_EX_TRANSPARENT`) — critical so the overlay doesn't interfere with normal app use. Only becomes interactive in edit mode.
- **One `.talon` file per process** — keeps generated files organized and easy to manage.
