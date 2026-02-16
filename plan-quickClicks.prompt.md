## Plan: Quick Clicks — Named Clickable Region Overlays (with click-type support)

**TL;DR**: Add a "Quick Clicks" system that lets users define named clickable regions per app (and optionally per window title). Regions are stored in `quick_clicks.json`, auto-shown when their associated app gains focus, hideable/showable via voice, and editable in-place. Each region has a configurable click type (single-left, double-left, right, etc.). Saying a region's name performs a single left-click by default; explicit phrases like "double click like" or "right click like" request other click types. The overlay also records the monitor resolution used when the profile was created — profiles are applied only to monitors with the same resolution and a warning is shown when the focused window is on a different-sized monitor. Multiple overlays (profiles) for the same application/process are supported so you can keep different overlays for different monitor sizes. The system auto-generates `.talon` files for voice integration. A foreground-window monitor (`SetWinEventHook`) in `ListenModeApplicationContext` drives auto-show/hide. The overlay reuses the transparent fullscreen WinForms pattern from `UIElementOverlayForm`. An in-place edit mode allows dragging regions, resizing via buttons, selecting click type, and creating new regions from the current mouse position.

**Steps**

### 1. Data Model — `Models/QuickClickRegion.cs` (new file)
- Define a `QuickClickProfile` class: `ProcessName` (string), `WindowTitlePattern` (string?, nullable for "any title"), `MonitorWidth` (int? — recorded monitor width in pixels; null = any monitor), `MonitorHeight` (int? — recorded monitor height in pixels; null = any monitor), `Regions` (list of `QuickClickRegion`).
- Define a `QuickClickRegion` class: `Name` (string), `X` (int), `Y` (int), `Width` (int, default 100), `Height` (int), `ClickType` (`QuickClickClickType`, default `Left`), `VoiceCommand` (string? — defaults to `Name` lowered).
- Add `public enum QuickClickClickType { Left, Double, Right, Middle }` (serialized as string in JSON).
- Coordinates are **relative to the app window's client area** (not absolute screen coords) so they survive window moves/resizes. Store as **pixel offsets from the window top-left** since most apps are used at consistent sizes, and percentages would be hard to reason about visually. We'll document that users should re-adjust after major window resizes.

### 2. Storage — `quick_clicks.json` + `Helpers/QuickClickLoader.cs` (new files)
- JSON file at the app directory, array of `QuickClickProfile` objects. Example (note `ClickType`):
  ```json
  [
    {
      "ProcessName": "chrome",
      "WindowTitlePattern": "YouTube",
      "MonitorWidth": 1920,
      "MonitorHeight": 1080,
      "Regions": [
        { "Name": "Like", "X": 450, "Y": 620, "Width": 100, "Height": 40, "ClickType": "Left" }
      ]
    }
  ]
  ```
- `QuickClickLoader` — static class modeled after `Helpers/MultiActionLoader.cs`. Methods: `Load()`, `Save(profiles)`, `GetProfilesForApp(processName, windowTitle, int? monitorWidth = null, int? monitorHeight = null)` which returns matching profiles (exact process match, optional title substring/regex match) and supports optional monitor-resolution filtering. Multiple profiles may exist for the same process/title but different recorded monitor resolutions. Loaded at startup in `NaturalLanguageInterpreter` static constructor alongside `MultiActionLoader.Load()`.

### 3. Action Model — add to `ActionModels.cs`
- Add `public enum QuickClickClickType { Left, Double, Right, Middle }` (if not added in the models file).
- Replace `ClickQuickClickAction` with a typed version: `public record ClickQuickClickAction(string RegionName, QuickClickClickType ClickType = QuickClickClickType.Left) : ActionBase;` — `ClickType` is optional; if omitted the action defaults to `Left` (saying the region name alone triggers a left click). The region's configured `ClickType` is retained for the editor and new-region defaults but does not affect bare-name voice invocation.
- Add `public record ShowQuickClicksAction() : ActionBase;`
- Add `public record HideQuickClicksAction() : ActionBase;`
- Add `public record EditQuickClicksAction() : ActionBase;`
- Add `public record CreateQuickClickAction() : ActionBase;` — for creating from current mouse position.

### 4. Foreground Window Monitor — `Helpers/ForegroundMonitor.cs` (new file)
- Use `SetWinEventHook(EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND, IntPtr.Zero, callback, 0, 0, WINEVENT_OUTOFCONTEXT)` to detect focus changes.
- Add required P/Invoke signatures: `SetWinEventHook`, `UnhookWinEvent`, and public `GetForegroundWindow` / `GetWindowText` to `Helpers/Win32ApiHelper.cs`.
- On each focus change: get process name via `CurrentApplicationHelper.GetCurrentProcessName( pilot checked)`, get window title via new `CurrentApplicationHelper.GetCurrentWindowTitle()`, determine the monitor resolution that contains the focused window (use `MonitorFromWindow` + `GetMonitorInfo` or `Screen.FromHandle`). Call `QuickClickLoader.GetProfilesForApp(processName, windowTitle, monitorWidth, monitorHeight)` to locate profiles recorded for that monitor resolution. If an exact-resolution profile exists → show `QuickClickOverlayForm` for that profile. If no exact-resolution profile exists but there are profiles for the same process on other resolutions → show a small non-interactive warning overlay explaining the mismatch and offer actions to create/apply a profile for the current monitor or open the editor; do **not** apply a mismatched profile. If no profiles at all → hide the overlay.
- Debounce rapid focus changes (100-200ms) using `DebounceGate` from `Helpers/DebounceGate.cs` or a simple timer.
- **Critical**: The overlay itself must NOT steal focus (use `WS_EX_NOACTIVATE` + `WS_EX_TOOLWINDOW` extended styles), otherwise it will trigger its own focus-change event in an infinite loop.
- Initialize in `ListenModeApplicationContext` constructor after hotkey registration (~line 87). Unhook in `ExitThreadCore()`.

### 5. Extend `CurrentApplicationHelper` — `CurrentApplicationHelper.cs`
- Add `public static string? GetCurrentWindowTitle()` — `GetForegroundWindow()` → `GetWindowText()` → return the title string. Reuse the `GetWindowText` P/Invoke already in `WindowFocusHelper` but make it accessible here.

### 6. Overlay Form — `QuickClickOverlayForm.cs` (new file)
- **Display mode** (non-interactive overlay):
  - Transparent fullscreen WinForms, same pattern as `UIElementOverlayForm.cs` (`TransparencyKey = Color.Lime`, borderless, topmost).
  - **Must be click-through** in display mode using `WS_EX_TRANSPARENT | WS_EX_LAYERED | WS_EX_NOACTIVATE | WS_EX_TOOLWINDOW` so the overlay doesn't block the underlying app.
  - Draw each region as a semi-transparent colored rectangle with the region name and a small `click-type` badge (e.g., "L", "2×", "R") as a label using GDI+.
  - Convert window-relative coordinates to screen coordinates using the target window's position (`GetWindowRect`).
  - **Monitor-resolution check**: only draw and activate regions when the current monitor's resolution matches the profile's recorded `MonitorWidth`/`MonitorHeight` (or when the profile's monitor fields are null meaning "any monitor"). If a profile exists for the app but does not match the current monitor resolution, do **not** apply the regions; instead render a prominent warning banner describing the mismatch and offer quick actions ("Create overlay for this monitor", "Open editor").
  - Singleton pattern (like `AutoClickOverlayForm`) with static `Show(profiles, windowHandle, monitorWidth, monitorHeight)`, `Hide()`, `IsVisible` methods. Thread-safe via `SynchronizationContext`.
  - Support multiple overlays per app/process (one profile per monitor resolution). When multiple profiles exist for the same process/title, show the profile that matches the monitor resolution.
  - Refresh/reposition when the target app window moves (poll via timer at ~200ms or hook `EVENT_OBJECT_LOCATIONCHANGE`).

- **Edit mode** (interactive):
  - Toggled by voice command ("edit quick clicks") or a small gear button drawn on the overlay.
  - Switch from click-through to interactive: remove `WS_EX_TRANSPARENT` flag, set `WS_EX_NOACTIVATE` to keep focus behavior correct.
  - Each region becomes **draggable** — mouse down on a region starts drag, mouse move repositions it, mouse up commits.
  - **Selection**: clicking a region selects it (highlighted border). Selected region shows **resize buttons** and a **click-type selector** rendered as small button-like rectangles (`[Left]` `[Double]` `[Right]`). Clicking a click-type button changes the region's `ClickType` immediately.
  - **New region from mouse**: a `[+ New]` button on the overlay. Clicking it enters "placement mode" — next click on the overlay captures mouse position, prompts for a name and default click type (small popup), creates a region at that point.
  - **Rectangle area mode**: a `[+ Rectangle]` button. First click sets corner 1, second click sets corner 2, then prompt for name and click type. This defines width/height from the two corners.
  - **Delete**: a `[Delete]` button visible when a region is selected.
  - **Done/Save**: a `[Save & Exit]` button that saves to `quick_clicks.json` via `QuickClickLoader.Save()` and switches back to display mode.
  - All edit UI rendered with GDI+ (no child WinForms controls, to keep the transparent overlay clean). Hit detection via manual coordinate checks in mouse event handlers.

### 7. Command Routing — `NaturalLanguageInterpreter.cs`
- In `InterpretAsync`, add a block (before the AI fallback, after emoji/general commands):
  - "show quick clicks" → `ShowQuickClicksAction`
  - "hide quick clicks" → `HideQuickClicksAction`
  - "edit quick clicks" → `EditQuickClicksAction`
  - "create quick click" → `CreateQuickClickAction`
  - **Click-type parsing**: recognize prefixes `double click`, `right click`, `click` (and synonyms). If a user says `double click <region>`, dispatch `ClickQuickClickAction(regionName, QuickClickClickType.Double)`.
  - **Dynamic region name matching**: load current profiles for the focused app, check if `text` matches any region's `Name` (case-insensitive, normalized). If matched with no explicit click-type prefix, dispatch `ClickQuickClickAction(regionName)` (the action defaults to `Left`) so the bare-name voice command performs a left click.
- In `ExecuteActionAsync` (same file, ~line 1161), add handlers:
  - `ClickQuickClickAction` → lookup region in current app's profile **for the monitor resolution that contains the focused window**; if no matching profile exists for the current monitor and `MatchOverlayByResolution` is enabled, show a warning and do **not** perform the click. Otherwise determine effective ClickType (action.ClickType if provided, otherwise `Left` — bare-name voice command defaults to `Left`); calculate screen center coordinates from window position + region offset; perform:
    - Left → SetCursorPos + mouse_event(LEFTDOWN/LEFTUP)
    - Double → two Left click sequences with a short gap (~30–80 ms)
    - Right → SetCursorPos + mouse_event(RIGHTDOWN/RIGHTUP)
    - Middle → similar if supported
    Reuse existing mouse helper methods (`AutoClickManager` / `Win32ApiHelper`) where possible.
  - `ShowQuickClicksAction` → `QuickClickOverlayForm.Show(...)`.
  - `HideQuickClicksAction` → `QuickClickOverlayForm.Hide()`.
  - `EditQuickClicksAction` → `QuickClickOverlayForm.EnterEditMode()`.
  - `CreateQuickClickAction` → get current cursor position, show `CreateQuickClickForm` (now includes click-type selector), create a region with selected click type, add to profile, save, refresh overlay.

### 8. Create Quick Click from Mouse Position — `CreateQuickClickForm.cs` (new, small)
- Minimal input dialog: a textbox for the region name, a click-type dropdown (`Left`/`Double`/`Right`), an option `Apply to this monitor only` (checked by default), + OK/Cancel buttons.
- Shown centered on screen when "create quick click" command is executed. The form captures the current monitor resolution and, if `Apply to this monitor only` is checked, the saved profile will record `MonitorWidth`/`MonitorHeight` for the current monitor; otherwise the profile's monitor fields are left null (apply to any monitor).
- After OK: creates a `QuickClickRegion` at the saved cursor position (captured *before* showing the dialog), with the selected `ClickType`; adds to the matching profile (or creates a new profile for the current app and monitor resolution), saves, and refreshes the overlay.

### 9. Talon File Generation — `Helpers/QuickClickTalonGenerator.cs` (new file)
- After any save to `quick_clicks.json`, auto-generate `.talon` files.
- One file per unique `ProcessName`: e.g., `quick_clicks_chrome.talon`, `quick_clicks_devenv.talon`.
- File placed in a configurable Talon user directory (add `TalonUserDirectory` setting to `Models/AppSettings.cs`) or alongside the app in a `talon_output/` folder.
- For each region, generate explicit voice phrases for supported click types plus the bare name. Example for region `Like`:
  ```
  # Auto-generated by NaturalCommands Quick Clicks — do not edit manually
  ^like$:
      user.run_application_csharp_natural("like")
  ^click like$:
      user.run_application_csharp_natural("click like")
  ^double click like$:
      user.run_application_csharp_natural("double click like")
  ^right click like$:
      user.run_application_csharp_natural("right click like")
  ```
  - `^like$` invokes a `Left` click (bare-name always performs a left click); explicit phrases override it.
- Regenerate fully on each save (overwrite).

### 10. Settings Integration — `Models/AppSettings.cs` + `SettingsForm.cs`
- Add a `QuickClickSettings` section to `AppSettings`:
  - `Enabled` (bool, default true)
  - `AutoShowOnFocus` (bool, default true)
  - `OverlayOpacity` (float, 0.0-1.0, default 0.5)
  - `DefaultRegionSize` (int, default 100)
  - `DefaultClickType` (`QuickClickClickType`, default `Left`)
  - `MatchOverlayByResolution` (bool, default true) — when true, profiles recorded for a specific monitor resolution will only apply on monitors with that exact resolution; when false, stored overlays may be applied across monitors (not recommended for pixel-accurate regions).
  - `TalonOutputDirectory` (string?, nullable — if null, output to app directory)
  
  Note: `DefaultClickType` affects newly-created regions and UI defaults; the bare-name voice command will still perform a left click. `MatchOverlayByResolution` controls the new resolution-matching behavior.
- Add a "Quick Clicks" tab to `SettingsForm` with controls for these settings (including default click type selector and the resolution-matching toggle).

### 11. Tray Menu Integration — `ListenModeApplicationContext.cs`
- Add a "Quick Clicks" submenu to the tray context menu with items:
  - "Show Quick Clicks" → triggers `ShowQuickClicksAction`
  - "Edit Quick Clicks" → triggers `EditQuickClicksAction`
  - "Create Quick Click at Mouse" → triggers `CreateQuickClickAction`
  - "Open quick_clicks.json" → opens the JSON file in default editor

### 12. New Files Summary

| File | Purpose |
|---|---|
| `Models/QuickClickRegion.cs` | Data model classes + `QuickClickClickType` |
| `Helpers/QuickClickLoader.cs` | JSON load/save |
| `Helpers/ForegroundMonitor.cs` | `SetWinEventHook` focus tracking |
| `QuickClickOverlayForm.cs` | Display + edit overlay (shows click-type badges & editor selectors) |
| `CreateQuickClickForm.cs` | Name + click-type input dialog |
| `Helpers/QuickClickTalonGenerator.cs` | `.talon` file generation (generates click/double/right variants) |

### 13. Modified Files Summary

| File | Changes |
|---|---|
| `ActionModels.cs` | Add `QuickClickClickType` enum and typed `ClickQuickClickAction` (region + optional click type) |
| `NaturalLanguageInterpreter.cs` | Add command matching for explicit click-type phrases + execution handlers |
| `CurrentApplicationHelper.cs` | Add `GetCurrentWindowTitle()` |
| `Helpers/Win32ApiHelper.cs` | Add `SetWinEventHook`, `UnhookWinEvent`, `GetForegroundWindow`, `GetWindowText` P/Invokes |
| `ListenModeApplicationContext.cs` | Initialize `ForegroundMonitor`, add tray menu items |
| `Models/AppSettings.cs` | Add `QuickClickSettings` (includes `DefaultClickType`) |
| `SettingsForm.cs` | Add Quick Clicks settings tab (click-type selector) |
| `Helpers/QuickClickTalonGenerator.cs` | Generate `.talon` phrases for `click`, `double click`, `right click`, and bare-name |

### Verification (additions for click types)
1. Build with `dotnet build NaturalCommands.csproj` — should compile clean.
2. Run in listen mode (`/listen`), open Chrome → verify overlay auto-appears if a chrome profile exists in `quick_clicks.json`.
3. Say "edit quick clicks" → verify edit mode activates with draggable regions, resize buttons, and click-type selector.
4. Create a region with `Double` click type → save → say "double click <region>" → verify a double-click occurs at the region center.
5. Create a region with `Right` click type → say the region's name (or "right click <region>") → verify right-click occurred (context menu opens).
6. Says region name without prefix → executes a single left click (bare-name always performs a left click).
7. Click `[+ New]` → click on overlay → enter name + choose click type → verify region created and saved.
8. Say "hide quick clicks" → verify overlay hides. Say "show quick clicks" → verify it returns.
9. Check that `.talon` files are generated with explicit `click`/`double click`/`right click` phrases and the bare name mapping.
10. Test with window title filtering — define a profile for `chrome` + `YouTube` title, verify overlay only shows on YouTube tabs.
11. Monitor-resolution verification — create two profiles for the same app: one recorded at 1920×1080 and another at 1366×768. Place the app on the 1920×1080 monitor and verify the 1920×1080 overlay appears. Move the app to the 1366×768 monitor and verify the 1366×768 overlay appears. If the app is on a monitor without a matching profile, verify a clear warning is shown and no mismatched overlay is applied.

### Design Decisions (click-type related)
- Bare region-name always performs a `Left` click for predictability. Per-region `ClickType` is retained as an editor/default and for explicit click-type phrases (e.g., "double click <region>").
- Monitor-specific profiles store the monitor resolution so regions remain pixel-accurate; this prevents mis-clicks on different-sized monitors. Supporting multiple profiles per app/process lets users keep separate overlays for laptop displays and external monitors.
- Explicit voice prefixes (`double click`, `right click`, `click`) always override the region default.
- We include `Right`/`Middle` click support for completeness, but primary focus is on `Left` (single) and `Double` as requested.
- Talon generator emits explicit variants so users can say either the bare name or the explicit phrase.
- Edit UI exposes click-type selection directly next to region controls for quick adjustments.

---

## Implementation tasks — checklist 📌
Use this checklist to track engineering progress. Update the `Status` column here and the workspace todo list (recommended) so both remain in sync.

| ID | Task | Files involved | Status | Notes |
|---:|---|---|---|---|
| 1 | Update plan (spec) | `plan-quickClicks.prompt.md` | completed | Plan updated with click-type & resolution support ✅ |
| 2 | Add data model | `Models/QuickClickRegion.cs` | completed | Profile + Region + `QuickClickClickType` enum |
| 3 | Implement loader | `Helpers/QuickClickLoader.cs` | completed ✅ | JSON load/save + GetProfilesForApp(...) |
| 4 | Talon generator | `Helpers/QuickClickTalonGenerator.cs` | completed ✅ | `.talon` file generation per process (auto-run on save) |
| 5 | Overlay UI scaffold | `QuickClickOverlayForm.cs` | completed ✅ | display + edit modes, selection, drag, basic editor UI |
| 6 | Create dialog | `CreateQuickClickForm.cs` | completed ✅ | name, click-type, monitor-scope option |
| 7 | Foreground monitor | `Helpers/ForegroundMonitor.cs` | completed ✅ | `SetWinEventHook` + debounce logic |
| 8 | Win32 P/Invokes | `Helpers/Win32ApiHelper.cs` | not-started | `SetWinEventHook`, `MonitorFromWindow`, `GetWindowText` |
| 9 | Current app helper | `CurrentApplicationHelper.cs` | not-started | `GetCurrentWindowTitle()` |
| 10 | Action model | `ActionModels.cs` | not-started | `ClickQuickClickAction`, `QuickClickClickType` |
| 11 | NLP + execution | `NaturalLanguageInterpreter.cs` | not-started | interpret + ExecuteActionAsync handlers |
| 12 | Tray + context menu | `ListenModeApplicationContext.cs` | not-started | add Quick Clicks submenu, init ForegroundMonitor |
| 13 | Settings UI & model | `Models/AppSettings.cs`, `SettingsForm.cs` | not-started | Quick Clicks settings tab, `MatchOverlayByResolution` |
| 14 | Tests — unit | `Tests/**` | not-started | loader, talon generator, action handlers |
| 15 | Tests — integration | `Tests/**` | not-started | overlay behavior, monitor-resolution logic |
| 16 | Docs & examples | `README.md`, `running-and-testing.md` | not-started | usage + JSON examples |
| 17 | QA & polish | n/a | not-started | accessibility, perf, UX tweaks |
| 18 | Release & talon deploy | n/a | not-started | generate + copy to `talon_output` or user dir |

### Recommended tracking
- Primary: maintain this checklist in the plan for quick reference.
- Live tracking: use the workspace todo list (created below) to mark items in-progress/completed — I can update those as I implement.
- Optional: create GitHub issues for major tasks (if you want PR/issue traceability).

---

### Next immediate task
- Implement the data model (`Models/QuickClickRegion.cs`) and the loader (`Helpers/QuickClickLoader.cs`).

> Note: tell me which task you want me to start first (I can mark it `in-progress` in the workspace todo list).
