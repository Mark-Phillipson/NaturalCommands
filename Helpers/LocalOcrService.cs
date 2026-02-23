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

            var tokens = normalized
                .ToLowerInvariant()
                .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Distinct()
                .ToArray();

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
                        int matchedTokenCount = tokens.Count(token => loweredLine.Contains(token));
                        if (matchedTokenCount == 0)
                        {
                            // Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: OCR line '{lineText}' - no tokens matched.");
                            continue;
                        }

                        Logger.LogDebug($"LocalOcrService.FindCandidatesAsync: OCR line '{lineText}' matched {matchedTokenCount}/{tokens.Length} tokens.");
                        var tokenCoverage = matchedTokenCount / (double)tokens.Length;
                        var confidence = 0.45 + (0.5 * tokenCoverage);
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
                        var rect = ToScreenRect(lineRect, virtualBounds);
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
                return candidates
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

        private static Rectangle ToScreenRect(Rect lineRect, Rectangle virtualBounds)
        {
            return new Rectangle(
                virtualBounds.Left + (int)Math.Round(lineRect.X),
                virtualBounds.Top + (int)Math.Round(lineRect.Y),
                Math.Max(1, (int)Math.Round(lineRect.Width)),
                Math.Max(1, (int)Math.Round(lineRect.Height)));
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

        // Backward compatibility wrapper - use FindCandidatesAsync instead
        [Obsolete("Use FindCandidatesAsync instead to avoid deadlocks")]
        public static List<VisualTargetCandidate> FindCandidates(string phrase)
        {
            return FindCandidatesAsync(phrase).GetAwaiter().GetResult();
        }
    }
}
