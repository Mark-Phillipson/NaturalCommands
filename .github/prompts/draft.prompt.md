
Title: Detect installed games and make them voice-launchable 🎮🔊

Description (short):
Add a feature that detects installed games on the user's system and ensures they are added to the app list so games can be launched by voice.

Acceptance criteria:
- The system scans common game sources (Steam library, Epic/GOG installers, Microsoft Store entries, Start Menu shortcuts, and common game folders) and returns a deduplicated list of installed games.
- Missing games are added (or flagged for confirmation) in the app/game catalog used by voice commands.
- Voice launch commands (e.g., "Play <game name>") can successfully launch detected games via the existing `AppLauncher` flow.
- Provide sensible logging and a small UI or settings entry to review/confirm detected matches.
- Include unit/integration tests where applicable and update docs (`README.md` or `running-and-testing.md`) with usage notes.

Implementation notes / tips (optional):
- Prefer Start Menu shortcuts and package manifests (MSIX) for reliability, and use Steam/Epic/GOG manifests or registry entries where available.
- Handle duplicates by normalizing titles and asking the user for confirmation in ambiguous cases.
- Add a background scan and a "Rescan" action in the UI.

Would you like me to also create an issue/PR with a more detailed implementation plan? 🔧