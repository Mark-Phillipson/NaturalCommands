using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Threading.Tasks;
using System.Windows.Forms;
using NaturalCommands.Models;
using Windows.Foundation;
using Windows.Globalization;
using Windows.Graphics.Imaging;
using Windows.Media.Ocr;

namespace NaturalCommands.Helpers
{
    public static class LocalOcrService
    {
        public static List<VisualTargetCandidate> FindCandidates(string phrase)
        {
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

                var virtualBounds = Screen.AllScreens.Aggregate(Rectangle.Empty, (current, screen) => Rectangle.Union(current, screen.Bounds));
                using var screenshot = new Bitmap(virtualBounds.Width, virtualBounds.Height);
                using (var graphics = Graphics.FromImage(screenshot))
                {
                    graphics.CopyFromScreen(virtualBounds.Left, virtualBounds.Top, 0, 0, virtualBounds.Size, CopyPixelOperation.SourceCopy);
                }

                var ocrResult = RunOcrAsync(screenshot).GetAwaiter().GetResult();
                if (ocrResult == null)
                {
                    return new List<VisualTargetCandidate>();
                }

                var candidates = new List<VisualTargetCandidate>();
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
                        continue;
                    }

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

                    var lineRect = BuildBoundingRectFromWords(line);
                    var rect = ToScreenRect(lineRect, virtualBounds);
                    if (rect.Width <= 0 || rect.Height <= 0)
                    {
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
                }

                return candidates
                    .OrderByDescending(c => c.Confidence)
                    .ThenByDescending(c => c.Bounds.Width * c.Bounds.Height)
                    .ToList();
            }
            catch (Exception ex)
            {
                Logger.LogError($"LocalOcrService failed: {ex.Message}");
                return new List<VisualTargetCandidate>();
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
            using var stream = new MemoryStream();
            bitmap.Save(stream, ImageFormat.Bmp);
            stream.Position = 0;

            using var randomAccessStream = stream.AsRandomAccessStream();
            var decoder = await BitmapDecoder.CreateAsync(randomAccessStream);
            var softwareBitmap = await decoder.GetSoftwareBitmapAsync(BitmapPixelFormat.Bgra8, BitmapAlphaMode.Premultiplied);

            var engine = OcrEngine.TryCreateFromUserProfileLanguages();
            if (engine == null)
            {
                engine = OcrEngine.TryCreateFromLanguage(new Language("en-US"));
            }

            if (engine == null)
            {
                Logger.LogWarning("LocalOcrService: no OCR engine available on this system.");
                return null;
            }

            return await engine.RecognizeAsync(softwareBitmap);
        }
    }
}
