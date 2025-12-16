# Plan: Voice Dictation Hotkey & Send Behavior 🎯

**TL;DR:** Add a resident hotkey (Windows+Ctrl+H) that opens the existing `VoiceDictationForm`. When voice typing stops, automatically focus the `Send` button (`btnSendCommand`) and set it as the form `AcceptButton` so the user can press **Enter** to send the command. Show the hotkey in the system tray and temporarily broadcast a “Listening… — press Enter to send” marquee message while the form is active.

---

## Goals

- Provide a keyboard shortcut to open the voice dictation UI (works even when Talon is running in background). 
- Make it easy to send the captured voice-typing as a command by focusing the Send button and enabling Enter to submit once dictation stops.
- Surface the shortcut in the tray so users discover it.
- Use the existing `VoiceDictationForm` and pipeline; avoid adding an external speech provider for the MVP.

## Steps (MVP)

1. Add `HotkeyRegistrar.cs` to register the default hotkey **Win + Ctrl + H** and raise activation events. 🔑
2. Add a resident `listen` mode in `Program.cs` (a tray `ApplicationContext`) to register the hotkey and manage lifecycle. 🖥️
3. On hotkey press, call the existing `VoiceDictationHelper.ShowVoiceDictation` to open `VoiceDictationForm` and start voice typing. 📣
4. In `VoiceDictationForm.cs`, add a `txtInput.TextChanged` debounce (700–1000ms). When the debounce fires, set `this.AcceptButton = btnSendCommand`, move focus to `btnSendCommand`, show a small transient label and play an optional stop tone. ⏱️↩️
5. Update `TrayNotificationHelper.cs` to include the hotkey string in the tray tooltip and context menu (e.g., “Open Voice Dictation (Win+Ctrl+H)”). 🧭
6. Implement `ShowTemporaryMarquee(string msg, int durationMs)` in `VoiceDictationForm` to temporarily replace marquee messages with “Listening — press Enter to send” while active. ➿

## Files & Symbols to change/add 🔧

- Add: `HotkeyRegistrar.cs` — wrapper for Win32 `RegisterHotKey`/`UnregisterHotKey` and a simple event.
- Modify: `Program.cs` — add `listen` resident mode and hook hotkey events into a tray `ApplicationContext`.
- Modify: `TrayNotificationHelper.cs` — add a menu item and tooltip showing the hotkey.
- Modify: `VoiceDictationForm.cs`
  - Add `txtInput.TextChanged` + debounce logic.
  - On debounce: set `this.AcceptButton = btnSendCommand`, call `btnSendCommand.Focus()`, show `lblTransient` message like “Listening finished — press Enter to send”, and (optionally) play a short sound.
  - Add `ShowTemporaryMarquee(string message, int durationMs)` to take over the marquee labels while the form is in active listening state.
- Tests: Add `Tests/VoiceDictationTests.cs` with unit tests for debounce behavior and `ShowTemporaryMarquee` behavior (mocking timers where necessary).

## UX details & behavior 🎯

- **Default hotkey:** Windows + Ctrl + H (configurable later). If the hotkey conflicts with other apps, provide a tray fallback and a small help note that the hotkey is in use and must be changed.
- **Visual feedback:** Use marquee takeover and `lblTransient` to show “Listening… — press Enter to send” while active and “Listening finished — press Enter to send” when finished. 
- **Audio feedback:** Optional short start and stop tones to increase confidence for users with auditory cues.
- **Talon compatibility:** The hotkey should attempt to open the form even if Talon runs; if Talon captures the same hotkey the app cannot override it — show a brief tray hint about choosing a different hotkey.

## Testing & rollout 🧪

- **Manual:** Start resident `listen` mode, press Win+Ctrl+H, verify form opens and auto-starts voice typing (Win+H), speak a phrase, verify that after speech the Send button receives focus and Enter sends the command through the existing pipeline. Test with Talon running.
- **Unit tests:** Debounce logic and `ShowTemporaryMarquee` behavior. Use mocking for timers and the helper that calls `VoiceDictationHelper.ShowVoiceDictation`.
- **Docs:** Update `running-and-testing.md` and README to document the hotkey and how to change it if conflicted.

## Follow-ups / Enhancements

- Make hotkey configurable from a settings UI.
- Add a small settings page to store and manage hotkey and audio cue preferences.
- Add stronger detection for Talon (e.g., check process name) and show an explicit option to force fallback when Talon is active (modifier key).

---

If you'd like, I can now implement the `HotkeyRegistrar` and a basic `listen` mode in `Program.cs` and add the debounce/focus behavior in `VoiceDictationForm.cs` as a small PR. Which part do you want me to start implementing first? ✅