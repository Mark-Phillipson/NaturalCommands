```instructions
---
applyTo: '**'
---

# Copilot Instructions — NaturalCommands skill

- **Primary goal:** Provide accurate, copy-pasteable commands for building, publishing, and running the app (including listen mode).
- **Verification:** Prefer manual verification via `bin/bin/app.log`. Do not run automated test suites to verify user changes.
- When assisting with publishing or startup shortcuts, point users to `.\scripts\publish-and-register-startup.ps1` and explain `-SelfContained` vs framework-dependent options.
- When answering, include examples and mention the resident hotkey and how to trigger listen mode: `dotnet run --framework net10.0-windows -- listen`.
- Keep answers short and actionable; include commands first, then a short explanation.

```