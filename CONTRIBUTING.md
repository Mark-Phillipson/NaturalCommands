Null-safety and defensive coding guidelines
========================================

This project prefers explicit, conservative null-handling so WinRT/COM and other
external APIs cannot cause silent crashes. Keep changes small and follow these
conventions when editing or adding C# files.

- Prefer safe navigation for chained access:
  - Use `?.` when accessing properties on values that may be null, e.g.
    `var binding = notification.Notification?.Visual?.GetBinding(...);`

- Declare locals nullable when assigning `null` or when the value may be absent:
  - `string? probeError = null;`

- Guard external or unstable objects before dereferencing:
  - Check `if (listener == null) { /* log and retry */ }` before calling methods.

- Avoid the null-forgiving operator (`!`) unless you have a provable invariant.

- When reading collections or maps, prefer `TryGetValue` and `string.IsNullOrWhiteSpace`.

- For interop or COM (WinRT) calls, always catch `COMException` and log details.

- CI guidance: run builds with warnings-as-errors to catch nullability regressions:
  ```powershell
  dotnet build -c Release -warnaserror
  ```

- Test-only mode: set the `NATURALCOMMANDS_TEST_MODE=1` environment variable or
  set `AppSettings.Instance.Behavior.TestMode = true` in a test to enable test-only
  mode. When enabled the app will avoid launching real terminals or other external
  programs (useful for CI and local unit tests).


If you have a specific interop scenario that requires a different approach, add a
brief comment at the call site explaining why the exception was made.

Thank you for keeping the codebase robust and maintainable.
