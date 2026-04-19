# Current Implementation Using Dotnet 

## Repository

GitHub: https://github.com/Mark-Phillipson/NaturalCommands

This repository contains `NaturalCommands`, a Windows automation utility that interprets natural-language commands and performs actions like clicking UI elements, controlling windows, typing text, and locating text on screen using OCR and UI Automation fallbacks. It includes helpers for visual targeting, quick clicks, voice dictation, notification handling, and integration with both local OCR (`Windows.Media.Ocr`) and OpenAI/cloud vision for advanced screen search.

### What the program does

- Listens for natural-language commands
- Finds UI elements by text or visual context
- Can click targets, show candidate overlays, or execute actions automatically
- Uses local OCR as a fallback when UI Automation can’t find a match
- Supports visual targeting settings, confidence thresholds, and cloud vision as optional behavior
- 
## OCR-based screen text location

### Key code files

- LocalOcrService.cs
- VisualTargetingService.cs

### How it works

1. `VisualTargetingService.IdentifyCandidates(...)` is the main entry point for visual text locating.
2. If UI Automation results are unavailable or low confidence, it falls back to OCR:
   - `VisualTargetingService.TryFromLocalOcr(...)`
   - which calls `LocalOcrService.FindCandidatesAsync(phrase)`

### Relevant code path

From VisualTargetingService.cs:

```csharp
if (shouldTryOcr)
{
    Logger.LogDebug($"VisualTargetingService: attempting OCR for phrase '{normalizedPhrase}'.");
    var ocrCandidates = TryFromLocalOcr(normalizedPhrase);
    Logger.LogDebug($"VisualTargetingService: OCR returned {ocrCandidates.Count} candidates.");
    ...
}
```

And `TryFromLocalOcr`:

```csharp
private static List<VisualTargetCandidate> TryFromLocalOcr(string phrase)
{
    var candidates = Task.Run(async () => await LocalOcrService.FindCandidatesAsync(phrase)).GetAwaiter().GetResult();
    return candidates;
}
```

### OCR implementation

From LocalOcrService.cs:

- Captures a screenshot of the foreground window
- Downscales large images for OCR stability
- Converts the screenshot to `SoftwareBitmap`
- Creates an OCR engine:
  - `OcrEngine.TryCreateFromUserProfileLanguages()`
  - fallback: `OcrEngine.TryCreateFromLanguage(new Language("en-US"))`
- Runs:
  - `engine.RecognizeAsync(softwareBitmap)`

### Library / software used

This code uses the Windows built-in OCR API, specifically:

- `Windows.Media.Ocr`
- `Windows.Graphics.Imaging`
- `Windows.Globalization`

So it is not using a third-party OCR library like Tesseract. It relies on the Windows Runtime OCR engine available on Windows 10/11 via the `Windows.Media.Ocr` namespace.

### Cross-platform / Talon Voice suitability

Because this implementation depends on the Windows Runtime OCR stack and Windows-specific screen capture/UI automation APIs, it is not directly portable to a cross-platform Python application or a Talon Voice extension.

For a cross-platform Talon Voice solution, a more appropriate approach would be:

- use a cross-platform OCR library such as Tesseract via Python bindings (`pytesseract`), or
- use a cloud OCR/vision API with a consistent cross-platform interface, and
- replace Windows-specific screen capture and UI automation with Talon/OS-agnostic capture mechanisms.

That said, the high-level architecture—capture screen data, run OCR, parse results, and map detected text to click targets—is still conceptually valid for a Talon Voice extension, but the concrete implementation details would need to change.

### Summary

- The OCR command logic is in LocalOcrService.cs
- It is invoked as a fallback from VisualTargetingService.cs
- The OCR engine is the Windows built-in `Windows.Media.Ocr` API, not a separate third-party OCR library