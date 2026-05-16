using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using NaturalCommands.Models;
using Windows.Foundation;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;
using Windows.Storage.Streams;

namespace NaturalCommands.Helpers
{
    public static class LocalOcrService
    {
        public static async Task<List<VisualTargetCandidate>> FindCandidatesAsync(string phrase, ScreenCaptureResult preCapture)
        {
            Bitmap? screenshot = null;
            try
            {
                var normalized = (phrase ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(normalized))
                {
                    return new List<VisualTargetCandidate>();
                }

                var tokens = BuildTokenListForPhrase(normalized);

                if (tokens.Length == 0)
                {
                    return new List<VisualTargetCandidate>();
                }

                Logger.LogDebug($"LocalOcrService.FindCandidatesAsync (preCapture): searching for '{phrase}' with tokens: {string.Join(", ", tokens)}.");

                OcrResult? ocrResult = null;
                Rectangle virtualBounds = Rectangle.Empty;
                
                try
                {
                    // Use provided capture bounds, falling back to foreground window if missing
                    virtualBounds = preCapture.VirtualBounds;
                    if (virtualBounds == Rectangle.Empty)
                    {
                        virtualBounds = WindowUtils.GetForegroundWindowBounds();
                    }

                    // Load screenshot bitmap from base64
                    var bytes = Convert.FromBase64String(preCapture.ImageBase64Jpeg);
                    using var ms = new MemoryStream(bytes);
                    screenshot = new Bitmap(ms);
                    Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: screenshot bitmap loaded, size={screenshot.Width}x{screenshot.Height}.");

                    // Downscale screenshot if it's too large (OCR can hang on very large images)
                    if (screenshot.Width > 2560 || screenshot.Height > 1440)
                    {
                        Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: screenshot is large ({screenshot.Width}x{screenshot.Height}), downscaling for OCR...");
                        double scaleX = 2560.0 / screenshot.Width;
                        double scaleY = 1440.0 / screenshot.Height;
                        double scale = Math.Min(scaleX, scaleY);
                        
                        int newWidth = (int)(screenshot.Width * scale);
                        int newHeight = (int)(screenshot.Height * scale);
                        Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: downscaling to {newWidth}x{newHeight} (scale={scale:0.00}).");
                        
                        var resized = new Bitmap(newWidth, newHeight);
                        using (var g = Graphics.FromImage(resized))
                        {
                            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                            g.DrawImage(screenshot, 0, 0, newWidth, newHeight);
                        }
                        screenshot.Dispose();
                        screenshot = resized;
                        Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: downscaled screenshot ready.");
                    }

                    Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: calling RunOcrAsync...");
                    try
                    {
                        ocrResult = await RunOcrAsync(screenshot);
                        Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: RunOcrAsync completed successfully.");
                    }
                    catch (Exception ocEx)
                    {
                        Logger.LogError($"LocalOcrService.FindCandidatesAsync: RunOcrAsync threw exception: {ocEx.GetType().Name}: {ocEx.Message}");
                        Logger.LogError($"LocalOcrService.FindCandidatesAsync: Stack trace: {ocEx.StackTrace}");
                        if (ocEx.InnerException != null)
                        {
                            Logger.LogError($"LocalOcrService.FindCandidatesAsync: Inner exception: {ocEx.InnerException.Message}");
                        }
                        throw;
                    }
                    
                    if (ocrResult == null)
                    {
                        Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: OCR returned null for phrase '{phrase}'.");
                        return new List<VisualTargetCandidate>();
                    }

                    Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: OCR found {ocrResult.Lines.Count} lines of text.");
                }
                catch (Exception ocEx)
                {
                    Logger.LogError($"LocalOcrService.FindCandidatesAsync: Screenshot/OCR error for phrase '{phrase}': {ocEx.GetType().Name}: {ocEx.Message}");
                    Logger.LogError($"LocalOcrService.FindCandidatesAsync: Full stack trace: {ocEx.StackTrace}");
                    return new List<VisualTargetCandidate>();
                }

                var candidates = new List<VisualTargetCandidate>();
                Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: starting to process {ocrResult.Lines.Count} OCR lines for phrase '{phrase}'...");
                // log each OCR line for debugging
                int lineIndex = 0;
                foreach (var debugLine in ocrResult.Lines)
                {
                    var dl = (debugLine.Text ?? string.Empty).Trim();
                    Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: OCR line #{lineIndex}: '{dl}'");
                    lineIndex++;
                }
                try
                {
                    foreach (var line in ocrResult.Lines)
                    {
                        var lineText = (line.Text ?? string.Empty).Trim();
                        if (string.IsNullOrWhiteSpace(lineText))
                        {
                            continue;
                        }

                        var loweredLine = lineText.ToLowerInvariant();

                        if (IsNoiseLine(lineText))
                        {
                            Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: skipping noise OCR line '{lineText}'");
                            continue;
                        }

                        int matchedTokenCount = tokens.Count(token => loweredLine.Contains(token) || (!string.IsNullOrEmpty(token) && lineText.Contains(token)));
                        if (matchedTokenCount == 0)
                        {
                            continue;
                        }

                        Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: OCR line '{lineText}' matched {matchedTokenCount}/{tokens.Length} tokens.");
                        var tokenCoverage = matchedTokenCount / (double)tokens.Length;
                        double confidence = 0.45 + (0.5 * tokenCoverage);
                        if (string.Equals(loweredLine, normalized.ToLowerInvariant(), StringComparison.OrdinalIgnoreCase))
                        {
                            confidence = 0.97;
                        }
                        else if (loweredLine.Contains(normalized, StringComparison.OrdinalIgnoreCase))
                        {
                            confidence = Math.Max(confidence, 0.85);
                        }

                        Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: computing bounding rect from {line.Words?.Count ?? 0} words...");
                        var lineRect = BuildBoundingRectFromWords(line);
                        Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: lineRect={lineRect.X},{lineRect.Y},{lineRect.Width}x{lineRect.Height}.");
                        var rect = ToScreenRect(lineRect, virtualBounds, screenshot.Size);
                        Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: screen rect={rect.X},{rect.Y},{rect.Width}x{rect.Height}.");
                        if (rect.Width <= 0 || rect.Height <= 0)
                        {
                            Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: rect has invalid dimensions, skipping.");
                            continue;
                        }

                        candidates.Add(new VisualTargetCandidate
                        {
                            Label = lineText,
                            Bounds = rect,
                            Confidence = Math.Max(0, Math.Min(1, confidence)),
                            Reason = $"OCR line match '{lineText}'",
                            Source = "ocr"
                        });
                        Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: added OCR candidate '{lineText}' with confidence {confidence:0.00} at ({rect.X},{rect.Y},{rect.Width}x{rect.Height}).");
                    }
                }
                catch (Exception loopEx)
                {
                    Logger.LogError($"LocalOcrService.FindCandidatesAsync: error in line processing loop: {loopEx.GetType().Name}: {loopEx.Message}");
                    Logger.LogError($"LocalOcrService.FindCandidatesAsync: loop error stack trace: {loopEx.StackTrace}");
                    throw;
                }

                Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: completed, found {candidates.Count} OCR candidates for '{phrase}'.");

                // Merge near-duplicate OCR candidates (same/similar label close together or overlapping)
                var merged = MergeOcrCandidates(candidates);
                Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: merged OCR candidates: {candidates.Count} -> {merged.Count}.");

                return merged
                    .OrderByDescending(c => c.Confidence)
                    .ThenByDescending(c => c.Bounds.Width * c.Bounds.Height)
                    .ToList();
            }
            catch (Exception ex)
            {
                Logger.LogError($"LocalOcrService.FindCandidatesAsync failed: {ex}");
                return new List<VisualTargetCandidate>();
            }
            finally
            {
                if (screenshot != null)
                {
                    screenshot.Dispose();
                }
            }
        }
public static async Task<List<VisualTargetCandidate>> FindCandidatesAsync(string phrase)
    {
        Bitmap? screenshot = null;
        try
        {
            var normalized = (phrase ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return new List<VisualTargetCandidate>();
            }

            var tokens = BuildTokenListForPhrase(normalized);

            if (tokens.Length == 0)
            {
                return new List<VisualTargetCandidate>();
            }

            Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: searching for '{phrase}' with tokens: {string.Join(", ", tokens)}.");

                OcrResult? ocrResult = null;
                Rectangle virtualBounds = Rectangle.Empty;
                
                try
                {
                    // Get only the foreground window bounds instead of all screens
                    virtualBounds = WindowUtils.GetForegroundWindowBounds();
                    
                    // Fallback to primary screen if we can't get foreground window
                    if (virtualBounds == Rectangle.Empty)
                    {
                        Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: Could not get foreground window bounds, falling back to primary screen.");
                        if (Screen.PrimaryScreen != null)
                        {
                            virtualBounds = Screen.PrimaryScreen.Bounds;
                        }
                        else if (Screen.AllScreens.Length > 0)
                        {
                            virtualBounds = Screen.AllScreens[0].Bounds;
                        }
                        else
                        {
                            virtualBounds = new Rectangle(0, 0, 1920, 1080); // Fallback default
                        }
                    }
                    
                    Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: virtualBounds={virtualBounds.Left},{virtualBounds.Top},{virtualBounds.Width}x{virtualBounds.Height}.");
                    
                    screenshot = new Bitmap(virtualBounds.Width, virtualBounds.Height);
                    Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: screenshot bitmap created, size={screenshot.Width}x{screenshot.Height}.");
                    
                    using (var graphics = Graphics.FromImage(screenshot))
                    {
                        graphics.CopyFromScreen(virtualBounds.Left, virtualBounds.Top, 0, 0, virtualBounds.Size, CopyPixelOperation.SourceCopy);
                    }
                    Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: screenshot captured successfully.");
                
                    // Downscale screenshot if it's too large (OCR can hang on very large images)
                    if (screenshot.Width > 2560 || screenshot.Height > 1440)
                    {
                        Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: screenshot is large ({screenshot.Width}x{screenshot.Height}), downscaling for OCR...");
                        double scaleX = 2560.0 / screenshot.Width;
                        double scaleY = 1440.0 / screenshot.Height;
                        double scale = Math.Min(scaleX, scaleY);
                        
                        int newWidth = (int)(screenshot.Width * scale);
                        int newHeight = (int)(screenshot.Height * scale);
                        Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: downscaling to {newWidth}x{newHeight} (scale={scale:0.00}).");
                        
                        var resized = new Bitmap(newWidth, newHeight);
                        using (var g = Graphics.FromImage(resized))
                        {
                            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                            g.DrawImage(screenshot, 0, 0, newWidth, newHeight);
                        }
                        screenshot.Dispose();
                        screenshot = resized;
                        Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: downscaled screenshot ready.");
                    }

                    Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: calling RunOcrAsync...");
                    try
                    {
                        ocrResult = await RunOcrAsync(screenshot);
                        Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: RunOcrAsync completed successfully.");
                    }
                    catch (Exception ocEx)
                    {
                        Logger.LogError($"LocalOcrService.FindCandidatesAsync: RunOcrAsync threw exception: {ocEx.GetType().Name}: {ocEx.Message}");
                        Logger.LogError($"LocalOcrService.FindCandidatesAsync: Stack trace: {ocEx.StackTrace}");
                        if (ocEx.InnerException != null)
                        {
                            Logger.LogError($"LocalOcrService.FindCandidatesAsync: Inner exception: {ocEx.InnerException.Message}");
                        }
                        throw;
                    }
                    
                    if (ocrResult == null)
                    {
                        Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: OCR returned null for phrase '{phrase}'.");
                        return new List<VisualTargetCandidate>();
                    }

                    Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: OCR found {ocrResult.Lines.Count} lines of text.");
                }
                catch (Exception ocEx)
                {
                    Logger.LogError($"LocalOcrService.FindCandidatesAsync: Screenshot/OCR error for phrase '{phrase}': {ocEx.GetType().Name}: {ocEx.Message}");
                    Logger.LogError($"LocalOcrService.FindCandidatesAsync: Full stack trace: {ocEx.StackTrace}");
                    return new List<VisualTargetCandidate>();
                }

                var candidates = new List<VisualTargetCandidate>();
                Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: starting to process {ocrResult.Lines.Count} OCR lines for phrase '{phrase}'...");
                // log each OCR line for debugging
                int lineIndex = 0;
                foreach (var debugLine in ocrResult.Lines)
                {
                    var dl = (debugLine.Text ?? string.Empty).Trim();
                    Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: OCR line #{lineIndex}: '{dl}'");
                    lineIndex++;
                }
                try
                {
                    foreach (var line in ocrResult.Lines)
                    {
                        var lineText = (line.Text ?? string.Empty).Trim();
                        if (string.IsNullOrWhiteSpace(lineText))
                        {
                            continue;
                        }

                        var loweredLine = lineText.ToLowerInvariant();

                        if (IsNoiseLine(lineText))
                        {
                            Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: skipping noise OCR line '{lineText}'");
                            continue;
                        }

                        int matchedTokenCount = tokens.Count(token => loweredLine.Contains(token) || (!string.IsNullOrEmpty(token) && lineText.Contains(token)));
                    if (matchedTokenCount == 0)
                    {
                        // Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: OCR line '{lineText}' - no tokens matched.");
                        continue;
                    }

                    Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: OCR line '{lineText}' matched {matchedTokenCount}/{tokens.Length} tokens.");
                    var tokenCoverage = matchedTokenCount / (double)tokens.Length;
                    double confidence = 0.45 + (0.5 * tokenCoverage);
                    if (string.Equals(loweredLine, normalized.ToLowerInvariant(), StringComparison.OrdinalIgnoreCase))
                    {
                        confidence = 0.97;
                    }
                    else if (loweredLine.Contains(normalized, StringComparison.OrdinalIgnoreCase))
                    {
                        confidence = Math.Max(confidence, 0.85);
                    }

                        Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: computing bounding rect from {line.Words?.Count ?? 0} words...");
                        var lineRect = BuildBoundingRectFromWords(line);
                        Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: lineRect={lineRect.X},{lineRect.Y},{lineRect.Width}x{lineRect.Height}.");
                        var rect = ToScreenRect(lineRect, virtualBounds, screenshot.Size);
                        Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: screen rect={rect.X},{rect.Y},{rect.Width}x{rect.Height}.");
                        if (rect.Width <= 0 || rect.Height <= 0)
                        {
                            Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: rect has invalid dimensions, skipping.");
                            continue;
                        }

                        candidates.Add(new VisualTargetCandidate
                        {
                            Label = lineText,
                            Bounds = rect,
                            Confidence = Math.Max(0, Math.Min(1, confidence)),
                            Reason = $"OCR line match '{lineText}'",
                            Source = "ocr"
                        });
                        Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: added OCR candidate '{lineText}' with confidence {confidence:0.00} at ({rect.X},{rect.Y},{rect.Width}x{rect.Height}).");
                    }
                }
                catch (Exception loopEx)
                {
                    Logger.LogError($"LocalOcrService.FindCandidatesAsync: error in line processing loop: {loopEx.GetType().Name}: {loopEx.Message}");
                    Logger.LogError($"LocalOcrService.FindCandidatesAsync: loop error stack trace: {loopEx.StackTrace}");
                    throw;
                }

                Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: completed, found {candidates.Count} OCR candidates for '{phrase}'.");
                // Merge near-duplicate OCR candidates before returning
                var merged = MergeOcrCandidates(candidates);
                Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: merged OCR candidates: {candidates.Count} -> {merged.Count}.");

                return merged
                    .OrderByDescending(c => c.Confidence)
                    .ThenByDescending(c => c.Bounds.Width * c.Bounds.Height)
                    .ToList();
            }
            catch (Exception ex)
            {
                Logger.LogError($"LocalOcrService.FindCandidatesAsync failed: {ex}");
                return new List<VisualTargetCandidate>();
            }
            finally
            {
                if (screenshot != null)
                {
                    screenshot.Dispose();
                }
            }
        }

        private static Rectangle ToScreenRect(Rect lineRect, Rectangle virtualBounds, System.Drawing.Size imageSize)
        {
            if (imageSize.Width <= 0 || imageSize.Height <= 0)
            {
                return new Rectangle(
                    virtualBounds.Left + (int)Math.Round(lineRect.X),
                    virtualBounds.Top + (int)Math.Round(lineRect.Y),
                    Math.Max(1, (int)Math.Round(lineRect.Width)),
                    Math.Max(1, (int)Math.Round(lineRect.Height)));
            }

            double scaleX = virtualBounds.Width / (double)imageSize.Width;
            double scaleY = virtualBounds.Height / (double)imageSize.Height;
            if (Math.Abs(scaleX - 1.0) > 0.001 || Math.Abs(scaleY - 1.0) > 0.001)
            {
                Logger.LogDebug($"ToScreenRect: imageSize={imageSize.Width}x{imageSize.Height}, virtualBounds={virtualBounds.Width}x{virtualBounds.Height}, scaleX={scaleX:0.000}, scaleY={scaleY:0.000}");
            }

            var x = virtualBounds.Left + (int)Math.Round(lineRect.X * scaleX);
            var y = virtualBounds.Top + (int)Math.Round(lineRect.Y * scaleY);
            var w = Math.Max(1, (int)Math.Round(lineRect.Width * scaleX));
            var h = Math.Max(1, (int)Math.Round(lineRect.Height * scaleY));

            return new Rectangle(x, y, w, h);
        }

        private static Rect BuildBoundingRectFromWords(OcrLine line)
        {
            if (line.Words == null || line.Words.Count == 0)
            {
                return new Rect(0, 0, 0, 0);
            }

            double minX = double.MaxValue;
            double minY = double.MaxValue;
            double maxX = 0;
            double maxY = 0;

            foreach (var word in line.Words)
            {
                var bounds = word.BoundingRect;
                minX = Math.Min(minX, bounds.X);
                minY = Math.Min(minY, bounds.Y);
                maxX = Math.Max(maxX, bounds.X + bounds.Width);
                maxY = Math.Max(maxY, bounds.Y + bounds.Height);
            }

            if (minX == double.MaxValue || minY == double.MaxValue)
            {
                return new Rect(0, 0, 0, 0);
            }

            return new Rect(minX, minY, Math.Max(0, maxX - minX), Math.Max(0, maxY - minY));
        }

        private static async Task<OcrResult?> RunOcrAsync(Bitmap bitmap)
        {
            try
            {
                Logger.LogDebug($"RunOcrAsync: starting, bitmap size={bitmap.Width}x{bitmap.Height}.");
                
                Logger.LogDebug($"RunOcrAsync: converting to SoftwareBitmap...");
                using var softwareBitmap = await ConvertToSoftwareBitmapAsync(bitmap);
                Logger.LogDebug($"RunOcrAsync: SoftwareBitmap conversion complete.");

                Logger.LogDebug($"RunOcrAsync: trying to create OCR engine from user profile languages...");
                var engine = OcrEngine.TryCreateFromUserProfileLanguages();
                if (engine == null)
                {
                    Logger.LogDebug($"RunOcrAsync: user profile languages engine failed, trying en-US...");
                    engine = OcrEngine.TryCreateFromLanguage(new Language("en-US"));
                }

                if (engine == null)
                {
                    Logger.LogWarning("RunOcrAsync: no OCR engine available on this system.");
                    return null;
                }

                Logger.LogDebug($"RunOcrAsync: OCR engine ready, starting recognition...");
                Logger.LogDebug($"RunOcrAsync: about to call engine.Recognize Async directly...");
                
                try
                {
                    Logger.LogDebug($"RunOcrAsync: calling RecognizeAsync directly on main thread...");
                    var result = await engine.RecognizeAsync(softwareBitmap);
                    Logger.LogDebug($"RunOcrAsync: RecognizeAsync returned, checking result...");
                    Logger.LogDebug($"RunOcrAsync: RecognizeAsync returned with {result?.Lines.Count ?? 0} lines.");
                    Logger.LogDebug($"RunOcrAsync: result is not null, returning...");
                    return result;
                }
                catch (ObjectDisposedException disposeEx)
                {
                    Logger.LogError($"RunOcrAsync: ObjectDisposedException during RecognizeAsync - resource was disposed: {disposeEx.Message}");
                    throw;
                }
                catch (Exception recEx)
                {
                    Logger.LogError($"RunOcrAsync: exception during RecognizeAsync: {recEx.GetType().Name}: {recEx.Message}");
                    Logger.LogError($"RunOcrAsync: exception stack trace: {recEx.StackTrace}");
                    throw;
                }
            }
            catch (Exception ex)
            {
                Logger.LogError($"RunOcrAsync failed: {ex.GetType().Name}: {ex.Message}");
                if (ex.InnerException != null)
                {
                    Logger.LogError($"RunOcrAsync inner exception: {ex.InnerException.GetType().Name}: {ex.InnerException.Message}");
                }
                Logger.LogError($"RunOcrAsync stack trace: {ex.StackTrace}");
                return null;
            }
        }

        private static async Task<SoftwareBitmap> ConvertToSoftwareBitmapAsync(Bitmap source)
        {
            try
            {
                Logger.LogDebug($"ConvertToSoftwareBitmapAsync: starting, source size={source.Width}x{source.Height}.");
                
                using var argb = new Bitmap(source.Width, source.Height, PixelFormat.Format32bppArgb);
                Logger.LogDebug($"ConvertToSoftwareBitmapAsync: ARGB bitmap created.");
                
                using (var graphics = Graphics.FromImage(argb))
                {
                    graphics.DrawImage(source, new Rectangle(0, 0, argb.Width, argb.Height));
                }
                Logger.LogDebug($"ConvertToSoftwareBitmapAsync: image drawn to ARGB bitmap.");

                using var memoryStream = new MemoryStream();
                argb.Save(memoryStream, ImageFormat.Bmp);
                Logger.LogDebug($"ConvertToSoftwareBitmapAsync: bitmap saved to memory stream, size={memoryStream.Length} bytes.");
                memoryStream.Position = 0;

                using var randomAccessStream = new InMemoryRandomAccessStream();
                using (var output = randomAccessStream.GetOutputStreamAt(0))
                {
                    using var writer = new DataWriter(output);
                    writer.WriteBytes(memoryStream.ToArray());
                    await writer.StoreAsync();
                    await writer.FlushAsync();
                    Logger.LogDebug($"ConvertToSoftwareBitmapAsync: written to InMemoryRandomAccessStream.");
                }

                randomAccessStream.Seek(0);
                Logger.LogDebug($"ConvertToSoftwareBitmapAsync: seeking to start of stream, creating decoder...");
                var decoder = await BitmapDecoder.CreateAsync(randomAccessStream);
                Logger.LogDebug($"ConvertToSoftwareBitmapAsync: decoder created, getting SoftwareBitmap...");
                var result = await decoder.GetSoftwareBitmapAsync(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied);
                Logger.LogDebug($"ConvertToSoftwareBitmapAsync: SoftwareBitmap created successfully.");
                return result;
            }
            catch (Exception ex)
            {
                Logger.LogError($"ConvertToSoftwareBitmapAsync failed: {ex.Message} - {ex.StackTrace}");
                throw;
            }
        }

        // card phrase helpers ------------------------------------------------
        // expose token construction logic for unit testing
        public static string[] BuildTokenListForPhrase(string phrase)
        {
            var normalized = (phrase ?? string.Empty).Trim();
            var punctuation = new[] { '.', ',', '!', '?', ';', ':', '"', '\'', '(', ')', '[', ']', '{', '}' };
            var tokenList = normalized
                .ToLowerInvariant()
                .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Select(token => token.Trim(punctuation))
                .Where(token => !string.IsNullOrWhiteSpace(token))
                .Distinct()
                .ToList();

            var requestedCard = TryParseCardPhrase(normalized);
            if (requestedCard != null)
            {
                string rankAbbrev = RankToAbbreviation(requestedCard.Value.rank);
                if (!string.IsNullOrWhiteSpace(rankAbbrev) && !tokenList.Contains(rankAbbrev))
                {
                    tokenList.Add(rankAbbrev);
                }

                string suitSymbol = SuitToSymbol(requestedCard.Value.suit);
                if (!string.IsNullOrWhiteSpace(suitSymbol) && !tokenList.Contains(suitSymbol))
                {
                    tokenList.Add(suitSymbol);
                }

                if (!tokenList.Contains(requestedCard.Value.suit))
                {
                    tokenList.Add(requestedCard.Value.suit);
                }
            }

            return tokenList.ToArray();
        }

        // Merge OCR candidates that are likely duplicates: same/similar label and
        // overlapping or very close bounding rectangles. This reduces noisy duplicate
        // hits (e.g., logo text + separate label boxes that refer to the same tile).
        private static List<VisualTargetCandidate> MergeOcrCandidates(List<VisualTargetCandidate> candidates)
        {
            if (candidates == null || candidates.Count <= 1)
                return candidates ?? new List<VisualTargetCandidate>();

            var used = new bool[candidates.Count];
            var groups = new System.Collections.Generic.List<System.Collections.Generic.List<int>>();

            for (int i = 0; i < candidates.Count; i++)
            {
                if (used[i]) continue;
                var g = new System.Collections.Generic.List<int> { i };
                used[i] = true;

                for (int j = i + 1; j < candidates.Count; j++)
                {
                    if (used[j]) continue;
                    if (ShouldMergeCandidates(candidates[i], candidates[j]))
                    {
                        g.Add(j);
                        used[j] = true;
                    }
                }

                groups.Add(g);
            }

            var result = new System.Collections.Generic.List<VisualTargetCandidate>();
            foreach (var g in groups)
            {
                if (g.Count == 1)
                {
                    result.Add(candidates[g[0]]);
                    continue;
                }

                // Merge group into a single candidate
                var union = candidates[g[0]].Bounds;
                double bestConfidence = candidates[g[0]].Confidence;
                string bestLabel = candidates[g[0]].Label ?? string.Empty;
                var reasons = new System.Text.StringBuilder();

                foreach (var idx in g)
                {
                    union = Rectangle.Union(union, candidates[idx].Bounds);
                    if (candidates[idx].Confidence > bestConfidence)
                    {
                        bestConfidence = candidates[idx].Confidence;
                        bestLabel = candidates[idx].Label ?? bestLabel;
                    }
                    if (!string.IsNullOrWhiteSpace(candidates[idx].Reason))
                    {
                        if (reasons.Length > 0) reasons.Append("; ");
                        reasons.Append(candidates[idx].Reason);
                    }
                }
                // debug output for merged group
                try
                {
                    var labels = string.Join(", ", g.Select(i => candidates[i].Label));
                    Logger.LogDebug($"MergeOcrCandidates: merging group indexes [{string.Join(",", g)}] labels [{labels}] into union {union}.");
                }
                catch { }
                result.Add(new VisualTargetCandidate
                {
                    Label = bestLabel,
                    Bounds = union,
                    Confidence = bestConfidence,
                    Reason = reasons.ToString(),
                    Source = "ocr"
                });
            }

            return result;
        }

        private static bool ShouldMergeCandidates(VisualTargetCandidate a, VisualTargetCandidate b)
        {
            if (a == null || b == null) return false;

            var la = NormalizeLabelForMerge(a.Label);
            var lb = NormalizeLabelForMerge(b.Label);
            if (string.IsNullOrEmpty(la) || string.IsNullOrEmpty(lb)) return false;

            bool labelsSimilar = la == lb || la.Contains(lb) || lb.Contains(la);

            // If rectangles intersect and labels are similar, definitely merge
            if (labelsSimilar && a.Bounds.IntersectsWith(b.Bounds)) return true;

            // Compute IoU (intersection over union) to detect strong overlap even if edges don't strictly intersect due to rounding
            try
            {
                var inter = Rectangle.Intersect(a.Bounds, b.Bounds);
                double interArea = (inter.Width > 0 && inter.Height > 0) ? (inter.Width * inter.Height) : 0;
                double unionArea = (a.Bounds.Width * a.Bounds.Height) + (b.Bounds.Width * b.Bounds.Height) - interArea;
                double iou = unionArea > 0 ? interArea / unionArea : 0;
                if (labelsSimilar && iou >= 0.20) return true;
            }
            catch { }

            // If labels are similar and centers are very close, merge
            if (labelsSimilar)
            {
                var ca = a.Center;
                var cb = b.Center;
                var dx = ca.X - cb.X;
                var dy = ca.Y - cb.Y;
                var distSq = dx * dx + dy * dy;
                // Use an averaged size-based hint but be less aggressive to avoid merging
                // distinct nearby tiles (reduce multiplier from 1.5 to 1.0 and clamp).
                var maxA = Math.Max(a.Bounds.Width, a.Bounds.Height);
                var maxB = Math.Max(b.Bounds.Width, b.Bounds.Height);
                var avgMax = (maxA + maxB) / 2.0;
                var sizeHint = Math.Max(24, (int)Math.Round(avgMax * 1.0));
                var maxSizeHint = 160;
                if (sizeHint > maxSizeHint) sizeHint = maxSizeHint;

                Logger.LogDebug($"ShouldMergeCandidates: labelsSimilar=true, centers distSq={distSq}, sizeHint={sizeHint}");

                if (distSq <= sizeHint * sizeHint) return true;
            }

            return false;
        }

        private static string NormalizeLabelForMerge(string label)
        {
            if (string.IsNullOrWhiteSpace(label)) return string.Empty;
            var cleaned = new System.Text.StringBuilder();
            foreach (var c in label)
            {
                if (char.IsLetterOrDigit(c) || char.IsWhiteSpace(c)) cleaned.Append(c);
            }
            return cleaned.ToString().ToLowerInvariant().Trim();
        }

        private static bool IsNoiseLine(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return true;
            var t = text.Trim();
            var lower = t.ToLowerInvariant();

            if (System.Text.RegularExpressions.Regex.IsMatch(t, "^\\[?(debug|info|error|warn|trace)\\b", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                return true;

            // Ignore command-line/log echo content that often includes the spoken phrase itself.
            if (lower.Contains("program.main")
                || lower.Contains("handlenaturalasync")
                || lower.Contains("interpreta")
                || lower.Contains("visualidentifyclickaction")
                || lower.Contains("dotnet run")
                || lower.Contains("dotnet build")
                || lower.Contains("get-content")
                || lower.Contains("remove-item")
                || lower.Contains("erroraction")
                || lower.Contains("cannot find path")
                || lower.Contains("ps c:\\")
                || lower.Contains("identify ")
                || lower.Contains(" -- natural ")
                || lower.Contains(" -tail "))
            {
                return true;
            }

            if (t.IndexOf("Generating patch", StringComparison.OrdinalIgnoreCase) >= 0
                || t.IndexOf("Edited", StringComparison.OrdinalIgnoreCase) >= 0
                || t.IndexOf("Read []", StringComparison.OrdinalIgnoreCase) >= 0
                || t.IndexOf("dotnet", StringComparison.OrdinalIgnoreCase) >= 0
                || t.IndexOf("Generating patch", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (t.IndexOf("C:\\", StringComparison.OrdinalIgnoreCase) >= 0
                || t.IndexOf("file://", StringComparison.OrdinalIgnoreCase) >= 0
                || t.IndexOf("://", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            int alnum = t.Count(char.IsLetterOrDigit);
            if (t.Length > 0 && (double)alnum / t.Length < 0.4) return true;
            if (t.Length <= 2 && t.All(char.IsDigit)) return true;

            return false;
        }

        private static (string rank, string suit)? TryParseCardPhrase(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            var regex = new System.Text.RegularExpressions.Regex(@"\b(?<rank>ace|king|queen|jack|10|[2-9]|two|three|four|five|six|seven|eight|nine|ten)\s+of\s+(?<suit>spades?|hearts?|diamonds?|clubs?)\b",
                System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.Compiled);
            var match = regex.Match(text);
            if (!match.Success)
            {
                return null;
            }

            var rank = NormalizeRank(match.Groups["rank"].Value);
            var suit = NormalizeSuit(match.Groups["suit"].Value);
            if (string.IsNullOrWhiteSpace(rank) || string.IsNullOrWhiteSpace(suit))
            {
                return null;
            }

            return (rank, suit);
        }

        private static string NormalizeRank(string rank)
        {
            return rank.Trim().ToLowerInvariant() switch
            {
                "2" or "two" => "two",
                "3" or "three" => "three",
                "4" or "four" => "four",
                "5" or "five" => "five",
                "6" or "six" => "six",
                "7" or "seven" => "seven",
                "8" or "eight" => "eight",
                "9" or "nine" => "nine",
                "10" or "ten" => "ten",
                "ace" => "ace",
                "king" => "king",
                "queen" => "queen",
                "jack" => "jack",
                _ => string.Empty
            };
        }

        private static string NormalizeSuit(string suit)
        {
            return suit.Trim().ToLowerInvariant() switch
            {
                "spade" or "spades" => "spades",
                "heart" or "hearts" => "hearts",
                "diamond" or "diamonds" => "diamonds",
                "club" or "clubs" => "clubs",
                _ => string.Empty
            };
        }

        private static string RankToAbbreviation(string rank)
        {
            return rank.ToLowerInvariant() switch
            {
                "ace" => "a",
                "king" => "k",
                "queen" => "q",
                "jack" => "j",
                "two" => "2",
                "three" => "3",
                "four" => "4",
                "five" => "5",
                "six" => "6",
                "seven" => "7",
                "eight" => "8",
                "nine" => "9",
                "ten" => "10",
                _ => string.Empty
            };
        }

        private static string SuitToSymbol(string suit)
        {
            return suit.ToLowerInvariant() switch
            {
                "spades" => "♠",
                "hearts" => "♥",
                "diamonds" => "♦",
                "clubs" => "♣",
                _ => string.Empty
            };
        }

        // Backward compatibility wrapper - use FindCandidatesAsync instead
        [Obsolete("Use FindCandidatesAsync instead to avoid deadlocks")]
        public static List<VisualTargetCandidate> FindCandidates(string phrase)
        {
            return FindCandidatesAsync(phrase).GetAwaiter().GetResult();
        }
    }
}
