using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using NaturalCommands.Models;

namespace NaturalCommands.Helpers
{
    public static class VisualTargetingService
    {
        private static readonly object _budgetLock = new();

        public static List<VisualTargetCandidate> IdentifyCandidates(string phrase)
        {
            var settings = AppSettings.Instance.VisualTargeting;
            var normalizedPhrase = (phrase ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalizedPhrase) || !settings.Enabled)
            {
                return new List<VisualTargetCandidate>();
            }

            var localUiCandidates = string.Equals(settings.FallbackMode, "ocr-only", StringComparison.OrdinalIgnoreCase)
                ? new List<VisualTargetCandidate>()
                : TryFromUiAutomation(normalizedPhrase);
            if (localUiCandidates.Count == 1
                && localUiCandidates[0].Confidence >= settings.AutoClickConfidenceThreshold
                && string.Equals(localUiCandidates[0].Source, "uia", StringComparison.OrdinalIgnoreCase))
            {
                Logger.LogDebug($"VisualTargetingService: using high-confidence UIA candidate without cloud call for '{normalizedPhrase}'.");
                return localUiCandidates
                    .OrderByDescending(c => c.Confidence)
                    .Take(Math.Max(1, settings.MaxCandidates))
                    .ToList();
            }

            List<VisualTargetCandidate> candidates = new();
            if (CanUseCloudVision(settings))
            {
                candidates = TryFromCloudVision(normalizedPhrase, settings);
            }

            if (candidates.Count == 0)
            {
                candidates = localUiCandidates;
            }

            if (candidates.Count == 0 && !string.Equals(settings.FallbackMode, "uia-only", StringComparison.OrdinalIgnoreCase))
            {
                candidates = TryFromLocalOcr(normalizedPhrase);
            }

            var maxCandidates = Math.Max(1, settings.MaxCandidates);
            return candidates
                .OrderByDescending(c => c.Confidence)
                .Take(maxCandidates)
                .ToList();
        }

        private static bool CanUseCloudVision(VisualTargetingSettings settings)
        {
            if (string.Equals(settings.ModelTier, "off", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (settings.MaxCloudCallsPerDay <= 0)
            {
                Logger.LogDebug("VisualTargetingService: cloud vision disabled by MaxCloudCallsPerDay=0.");
                return false;
            }

            var apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
            if (string.IsNullOrWhiteSpace(apiKey))
            {
                Logger.LogWarning("VisualTargetingService: OPENAI_API_KEY not set; skipping cloud vision.");
                return false;
            }

            lock (_budgetLock)
            {
                var today = DateTime.UtcNow.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                if (!string.Equals(settings.CloudUsageDateUtc, today, StringComparison.Ordinal))
                {
                    settings.CloudUsageDateUtc = today;
                    settings.CloudCallsUsedToday = 0;
                }

                if (settings.CloudCallsUsedToday >= settings.MaxCloudCallsPerDay)
                {
                    Logger.LogInfo($"VisualTargetingService: cloud vision budget reached for {today} ({settings.CloudCallsUsedToday}/{settings.MaxCloudCallsPerDay}).");
                    return false;
                }

                settings.CloudCallsUsedToday++;
                TrySaveSettings();
                return true;
            }
        }

        private static List<VisualTargetCandidate> TryFromCloudVision(string phrase, VisualTargetingSettings settings)
        {
            try
            {
                var capture = ScreenCaptureService.CaptureAllScreensJpeg(maxLongEdge: 1400, quality: 55);
                var model = MapModelTier(settings.ModelTier);
                var payloadJson = BuildCloudRequestJson(model, phrase, capture.ImageBase64Jpeg);

                using var client = new HttpClient();
                client.Timeout = TimeSpan.FromSeconds(9);
                var apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY") ?? string.Empty;
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

                using var request = new HttpRequestMessage(HttpMethod.Post, "https://api.openai.com/v1/chat/completions")
                {
                    Content = new StringContent(payloadJson, Encoding.UTF8, "application/json")
                };

                using var response = client.Send(request);
                var responseBody = response.Content.ReadAsStringAsync().GetAwaiter().GetResult();
                if (!response.IsSuccessStatusCode)
                {
                    Logger.LogWarning($"VisualTargetingService: cloud vision request failed ({(int)response.StatusCode}): {responseBody}");
                    return new List<VisualTargetCandidate>();
                }

                var content = ExtractMessageContent(responseBody);
                if (string.IsNullOrWhiteSpace(content))
                {
                    Logger.LogWarning("VisualTargetingService: cloud vision returned empty content.");
                    return new List<VisualTargetCandidate>();
                }

                var parsed = ParseCandidatesFromJson(content, capture);
                if (parsed.Count > 0)
                {
                    Logger.LogDebug($"VisualTargetingService: cloud vision returned {parsed.Count} candidates.");
                }
                return parsed;
            }
            catch (Exception ex)
            {
                Logger.LogError($"VisualTargetingService cloud vision failed: {ex.Message}");
                return new List<VisualTargetCandidate>();
            }
        }

        private static string MapModelTier(string tier)
        {
            if (string.Equals(tier, "fast", StringComparison.OrdinalIgnoreCase))
            {
                return "gpt-4.1-mini";
            }

            return "gpt-4.1";
        }

        private static string BuildCloudRequestJson(string model, string phrase, string jpegBase64)
        {
            var imageDataUri = $"data:image/jpeg;base64,{jpegBase64}";
            var payload = new
            {
                model,
                temperature = 0.0,
                max_tokens = 500,
                messages = new object[]
                {
                    new
                    {
                        role = "system",
                        content = "You identify UI targets in screenshots. Return only JSON with this shape: {\"candidates\":[{\"label\":string,\"bbox\":{\"x\":number,\"y\":number,\"width\":number,\"height\":number},\"confidence\":number,\"reason\":string}]}. Use screenshot pixel coordinates for bbox. Confidence must be 0..1. Return up to 8 candidates sorted best-first."
                    },
                    new
                    {
                        role = "user",
                        content = new object[]
                        {
                            new { type = "text", text = $"Find UI elements matching this phrase: '{phrase}'" },
                            new { type = "image_url", image_url = new { url = imageDataUri } }
                        }
                    }
                }
            };

            return JsonSerializer.Serialize(payload);
        }

        private static string ExtractMessageContent(string responseJson)
        {
            using var document = JsonDocument.Parse(responseJson);
            if (!document.RootElement.TryGetProperty("choices", out var choices) || choices.ValueKind != JsonValueKind.Array || choices.GetArrayLength() == 0)
            {
                return string.Empty;
            }

            var message = choices[0].GetProperty("message");
            if (!message.TryGetProperty("content", out var contentElement))
            {
                return string.Empty;
            }

            var content = contentElement.GetString() ?? string.Empty;
            content = content.Trim();
            if (content.StartsWith("```", StringComparison.Ordinal))
            {
                var start = content.IndexOf('\n');
                var end = content.LastIndexOf("```", StringComparison.Ordinal);
                if (start >= 0 && end > start)
                {
                    content = content.Substring(start + 1, end - start - 1).Trim();
                }
            }

            return content;
        }

        private static List<VisualTargetCandidate> ParseCandidatesFromJson(string json, ScreenCaptureResult capture)
        {
            var candidates = new List<VisualTargetCandidate>();
            using var document = JsonDocument.Parse(json);
            if (!document.RootElement.TryGetProperty("candidates", out var array) || array.ValueKind != JsonValueKind.Array)
            {
                return candidates;
            }

            double scaleX = capture.VirtualBounds.Width / (double)Math.Max(1, capture.Width);
            double scaleY = capture.VirtualBounds.Height / (double)Math.Max(1, capture.Height);

            foreach (var item in array.EnumerateArray())
            {
                if (!item.TryGetProperty("bbox", out var bbox) || bbox.ValueKind != JsonValueKind.Object)
                {
                    continue;
                }

                double x = bbox.TryGetProperty("x", out var xEl) && xEl.TryGetDouble(out var xv) ? xv : 0;
                double y = bbox.TryGetProperty("y", out var yEl) && yEl.TryGetDouble(out var yv) ? yv : 0;
                double width = bbox.TryGetProperty("width", out var wEl) && wEl.TryGetDouble(out var wv) ? wv : 0;
                double height = bbox.TryGetProperty("height", out var hEl) && hEl.TryGetDouble(out var hv) ? hv : 0;

                if (x <= 1 && y <= 1 && width <= 1 && height <= 1)
                {
                    x *= capture.Width;
                    y *= capture.Height;
                    width *= capture.Width;
                    height *= capture.Height;
                }

                var mappedRect = new System.Drawing.Rectangle(
                    capture.VirtualBounds.Left + (int)Math.Round(x * scaleX),
                    capture.VirtualBounds.Top + (int)Math.Round(y * scaleY),
                    Math.Max(1, (int)Math.Round(width * scaleX)),
                    Math.Max(1, (int)Math.Round(height * scaleY)));

                var confidence = item.TryGetProperty("confidence", out var cEl) && cEl.TryGetDouble(out var cv) ? cv : 0.5;
                confidence = Math.Max(0, Math.Min(1, confidence));

                var label = item.TryGetProperty("label", out var lEl) ? (lEl.GetString() ?? "target") : "target";
                var reason = item.TryGetProperty("reason", out var rEl) ? (rEl.GetString() ?? string.Empty) : string.Empty;

                candidates.Add(new VisualTargetCandidate
                {
                    Label = label,
                    Bounds = mappedRect,
                    Confidence = confidence,
                    Reason = reason,
                    Source = "vision"
                });
            }

            return candidates;
        }

        private static void TrySaveSettings()
        {
            try
            {
                AppSettings.Instance.Save();
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"VisualTargetingService: failed to persist budget counters: {ex.Message}");
            }
        }

        private static List<VisualTargetCandidate> TryFromUiAutomation(string phrase)
        {
            try
            {
                var loweredPhrase = phrase.ToLowerInvariant();
                var tokens = loweredPhrase.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                var elements = UIAutomationHelper.EnumerateClickableElements(scopeToActiveWindow: false);

                var matches = new List<VisualTargetCandidate>();
                foreach (var element in elements)
                {
                    if (element.Bounds.Width <= 0 || element.Bounds.Height <= 0)
                    {
                        continue;
                    }

                    var name = (element.Name ?? string.Empty).Trim();
                    if (string.IsNullOrWhiteSpace(name))
                    {
                        continue;
                    }

                    var loweredName = name.ToLowerInvariant();
                    if (!tokens.All(t => loweredName.Contains(t)))
                    {
                        continue;
                    }

                    double score = 0.55;
                    if (string.Equals(loweredName, loweredPhrase, StringComparison.OrdinalIgnoreCase))
                    {
                        score = 0.95;
                    }
                    else if (loweredName.StartsWith(loweredPhrase, StringComparison.OrdinalIgnoreCase))
                    {
                        score = 0.88;
                    }
                    else if (loweredName.Contains(loweredPhrase, StringComparison.OrdinalIgnoreCase))
                    {
                        score = 0.8;
                    }

                    matches.Add(new VisualTargetCandidate
                    {
                        Label = name,
                        Bounds = element.Bounds,
                        Confidence = score,
                        Reason = $"UIA match '{name}'",
                        Source = "uia"
                    });
                }

                return matches;
            }
            catch (Exception ex)
            {
                Logger.LogError($"VisualTargetingService UIA fallback failed: {ex.Message}");
                return new List<VisualTargetCandidate>();
            }
        }

        private static List<VisualTargetCandidate> TryFromLocalOcr(string phrase)
        {
            try
            {
                var candidates = LocalOcrService.FindCandidates(phrase);
                if (candidates.Count > 0)
                {
                    Logger.LogDebug($"VisualTargetingService: OCR fallback returned {candidates.Count} candidates for '{phrase}'.");
                }
                else
                {
                    Logger.LogDebug($"VisualTargetingService: OCR fallback found no matches for '{phrase}'.");
                }
                return candidates;
            }
            catch (Exception ex)
            {
                Logger.LogError($"VisualTargetingService OCR fallback failed: {ex.Message}");
                return new List<VisualTargetCandidate>();
            }
        }
    }
}
