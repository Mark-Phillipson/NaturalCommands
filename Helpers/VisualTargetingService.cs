using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using NaturalCommands.Models;
using NaturalCommands;

namespace NaturalCommands.Helpers
{
    public static class VisualTargetingService
    {
        private static readonly object _budgetLock = new();
        private static readonly Regex CardConnectorMisrecognitionRegex = new(@"\b(?<rank>ace|king|queen|jack|10|[2-9]|two|three|four|five|six|seven|eight|nine|ten)\s+off\s+(?<suit>spades?|hearts?|diamonds?|clubs?)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex CardPhraseRegex = new(@"\b(?<rank>ace|king|queen|jack|10|[2-9]|two|three|four|five|six|seven|eight|nine|ten)\s+of\s+(?<suit>spades?|hearts?|diamonds?|clubs?)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static List<VisualTargetCandidate> IdentifyCandidates(string phrase)
        {
            var settings = AppSettings.Instance.VisualTargeting;
            // normalize early so all downstream logic sees the same phrase transformation
            var normalizedPhrase = NormalizePhrase(phrase);
            if (string.IsNullOrWhiteSpace(normalizedPhrase) || !settings.Enabled)
            {
                return new List<VisualTargetCandidate>();
            }

            var currentProcessName = NaturalCommands.CurrentApplicationHelper.GetCurrentProcessName() ?? string.Empty;
            Logger.LogInfo($"VisualTargetingService: IdentifyCandidates called - phrase='{phrase}', normalized='{normalizedPhrase}', currentProcess='{currentProcessName}'.");
            var foregroundBounds = WindowUtils.GetForegroundWindowBounds();
            if (foregroundBounds == System.Drawing.Rectangle.Empty)
            {
                Logger.LogWarning("VisualTargetingService: foreground window bounds unavailable; refusing off-window matching.");
                return new List<VisualTargetCandidate>();
            }
            
            // Capture a screenshot once and pass it through to vision/OCR paths
            ScreenCaptureResult? preCapture = null;
            try
            {
                preCapture = ScreenCaptureService.CaptureRegionJpeg(foregroundBounds, maxLongEdge: 1400, quality: 60);
                Logger.LogDebug($"VisualTargetingService: preCapture VirtualBounds={preCapture.VirtualBounds.Left},{preCapture.VirtualBounds.Top},{preCapture.VirtualBounds.Width}x{preCapture.VirtualBounds.Height}; image={preCapture.Width}x{preCapture.Height}.");
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"VisualTargetingService: pre-capture failed: {ex.Message}");
                return new List<VisualTargetCandidate>();
            }
            var preferOcrForCurrentApp = ShouldPreferOcrForApp(currentProcessName)
                && !string.Equals(settings.FallbackMode, "uia-only", StringComparison.OrdinalIgnoreCase);

            if (preferOcrForCurrentApp)
            {
                Logger.LogInfo($"VisualTargetingService: preferring OCR for process '{currentProcessName}'.");
            }

            var localUiCandidates = string.Equals(settings.FallbackMode, "ocr-only", StringComparison.OrdinalIgnoreCase) || preferOcrForCurrentApp
                ? new List<VisualTargetCandidate>()
                : TryFromUiAutomation(normalizedPhrase);
            localUiCandidates = FilterCandidatesToBounds(localUiCandidates, foregroundBounds, "UIA");
            Logger.LogDebug($"VisualTargetingService: localUiCandidates.Count={localUiCandidates.Count} for phrase '{normalizedPhrase}'.");
            if (localUiCandidates.Count > 0)
            {
                Logger.LogDebug($"VisualTargetingService: UIA candidates: {string.Join(", ", localUiCandidates.Select(c => $"'{c.Label}'({c.Confidence:0.00})"))}");
            }
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
                candidates = TryFromCloudVision(normalizedPhrase, settings, preCapture);
            }

            if (candidates.Count == 0)
            {
                candidates = localUiCandidates;
                Logger.LogDebug($"VisualTargetingService: cloud vision returned 0 candidates, using localUiCandidates.");
            }

            if (!string.Equals(settings.FallbackMode, "uia-only", StringComparison.OrdinalIgnoreCase))
            {
                var shouldTryOcr = candidates.Count == 0;
                Logger.LogDebug($"VisualTargetingService: initial candidates.Count={candidates.Count}, shouldTryOcr={shouldTryOcr}.");
                if (!shouldTryOcr && string.Equals(candidates[0].Source, "uia", StringComparison.OrdinalIgnoreCase) && candidates[0].Confidence < 0.86)
                {
                    shouldTryOcr = true;
                }

                if (shouldTryOcr)
                {
                    Logger.LogDebug($"VisualTargetingService: attempting OCR for phrase '{normalizedPhrase}'.");
                    var ocrCandidates = TryFromLocalOcr(normalizedPhrase, preCapture);
                    ocrCandidates = FilterCandidatesToBounds(ocrCandidates, foregroundBounds, "OCR");
                    Logger.LogDebug($"VisualTargetingService: OCR returned {ocrCandidates.Count} candidates.");
                    if (candidates.Count == 0)
                    {
                        candidates = ocrCandidates;
                        Logger.LogDebug($"VisualTargetingService: using OCR candidates (was 0 before).");
                    }
                    else if (ocrCandidates.Count > 0)
                    {
                        candidates = candidates
                            .Concat(ocrCandidates)
                            .OrderByDescending(c => c.Confidence)
                            .ToList();
                    }
                }
            }

            candidates = FilterCandidatesToBounds(candidates, foregroundBounds, "final");

            var maxCandidates = Math.Max(1, settings.MaxCandidates);
            var finalResult = candidates
                .OrderByDescending(c => c.Confidence)
                .Take(maxCandidates)
                .ToList();
            if (finalResult.Count > 0)
            {
                Logger.LogDebug($"VisualTargetingService: returning {finalResult.Count} candidates; top is '{finalResult[0].Label}' (confidence {finalResult[0].Confidence:0.00}, source '{finalResult[0].Source}').");
            }
            else
            {
                Logger.LogDebug($"VisualTargetingService: returning 0 candidates for phrase '{normalizedPhrase}'.");
            }
            Logger.LogInfo($"VisualTargetingService.IdentifyCandidates complete: returning {finalResult.Count} candidates for '{normalizedPhrase}'.");
            return finalResult;
        }

        private static List<VisualTargetCandidate> FilterCandidatesToBounds(
            List<VisualTargetCandidate> candidates,
            System.Drawing.Rectangle foregroundBounds,
            string source)
        {
            if (candidates == null || candidates.Count == 0)
            {
                return new List<VisualTargetCandidate>();
            }

            var filtered = candidates
                .Where(c => c.Bounds.Width > 0 && c.Bounds.Height > 0 && foregroundBounds.IntersectsWith(c.Bounds))
                .ToList();

            if (filtered.Count != candidates.Count)
            {
                Logger.LogDebug($"VisualTargetingService: {source} bounds filter kept {filtered.Count}/{candidates.Count} candidates inside foreground window {foregroundBounds.Left},{foregroundBounds.Top},{foregroundBounds.Width}x{foregroundBounds.Height}.");
            }

            return filtered;
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

        private static List<VisualTargetCandidate> TryFromCloudVision(string phrase, VisualTargetingSettings settings, ScreenCaptureResult? preCapture = null)
        {
            try
            {
                var capture = preCapture ?? ScreenCaptureService.CaptureAllScreensJpeg(maxLongEdge: 1400, quality: 55);
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

        /// <summary>
        /// Public helper for diagnostics: capture screenshot and send to cloud vision API using current settings.
        /// Returns whatever candidates the model produces, or an empty list if cloud vision cannot be used.
        /// </summary>
        public static List<VisualTargetCandidate> GetCloudVisionCandidates(string phrase)
        {
            var settings = AppSettings.Instance.VisualTargeting;
            if (!CanUseCloudVision(settings))
            {
                Logger.LogInfo("VisualTargetingService: GetCloudVisionCandidates skipped because cloud vision is unavailable.");
                return new List<VisualTargetCandidate>();
            }

            var normalized = NormalizeCardConnectorMisrecognitions((phrase ?? string.Empty).Trim());
            Logger.LogInfo($"VisualTargetingService: GetCloudVisionCandidates invoked for phrase '{normalized}'");

            // capture screenshot ourselves so we can save it for debugging
            var capture = ScreenCaptureService.CaptureAllScreensJpeg(maxLongEdge: 1400, quality: 60);
            try
            {
                var tmp = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
                    "nc-screenshot-" + DateTime.UtcNow.ToString("yyyyMMdd-HHmmss") + ".jpg");
                System.IO.File.WriteAllBytes(tmp, Convert.FromBase64String(capture.ImageBase64Jpeg));
                Logger.LogInfo($"VisualTargetingService: saved debug screenshot to {tmp}");
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"VisualTargetingService: failed to save debug screenshot: {ex.Message}");
            }

            // feed capture manually to cloud request to avoid recapturing with different parameters
            return TryFromCloudVision(normalized, settings, capture);
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
                var requestedCard = ParseCardFromPhrase(loweredPhrase);
                var requestedRankOnly = requestedCard == null ? ParseRankOnlyQuery(loweredPhrase) : string.Empty;
                var elements = UIAutomationHelper.EnumerateClickableElements(scopeToActiveWindow: true);

                Logger.LogDebug($"TryFromUiAutomation: searching for phrase '{phrase}' (tokens: {string.Join(", ", tokens)}), found {elements.Count} clickable elements.");
                // also dump element names for troubleshooting
                if (elements.Count > 0)
                {
                    var sampleNames = elements
                        .Take(20)
                        .Select(e => {
                            var tc = (e.TextCandidates != null && e.TextCandidates.Count > 0) ? string.Join("|", e.TextCandidates.Take(3)) : "";
                            return string.IsNullOrWhiteSpace(e.Name) ? (string.IsNullOrWhiteSpace(tc) ? "<empty>" : $"[{tc}]") : (string.IsNullOrWhiteSpace(tc) ? e.Name : $"{e.Name} [{tc}]");
                        })
                        .ToList();
                    Logger.LogDebug($"TryFromUiAutomation: sample element names/candidates: {string.Join(", ", sampleNames)}");
                }

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

                    if (name.Length > 140 || name.Contains('\n') || name.Contains('\r'))
                    {
                        continue;
                    }

                    var loweredName = name.ToLowerInvariant();
                    var cardsInName = ParseCardsFromText(loweredName);
                    double score;
                    if (requestedCard != null)
                    {
                        var matchingCards = cardsInName.Where(c => c.Equals(requestedCard)).ToList();
                        if (matchingCards.Count == 0)
                        {
                            continue;
                        }

                        if (cardsInName.Count > 1)
                        {
                            matches.AddRange(BuildStackCardCandidates(
                                element.Bounds,
                                cardsInName,
                                card => card.Equals(requestedCard),
                                0.9,
                                "UIA stacked-card exact match"));
                            continue;
                        }

                        var normalizedCardPhrase = $"{requestedCard.Rank} of {requestedCard.Suit}";
                        if (string.Equals(loweredName, normalizedCardPhrase, StringComparison.OrdinalIgnoreCase))
                        {
                            score = 0.98;
                        }
                        else if (ContainsWholePhrase(loweredName, normalizedCardPhrase))
                        {
                            score = 0.9;
                        }
                        else
                        {
                            score = 0.82;
                        }
                    }
                    else
                    {
                        if (cardsInName.Count > 1 && !string.IsNullOrWhiteSpace(requestedRankOnly))
                        {
                            var stackRankMatches = BuildStackCardCandidates(
                                element.Bounds,
                                cardsInName,
                                card => string.Equals(card.Rank, requestedRankOnly, StringComparison.Ordinal),
                                0.74,
                                $"UIA stacked-card rank match '{requestedRankOnly}'");

                            if (stackRankMatches.Count > 0)
                            {
                                matches.AddRange(stackRankMatches);
                                continue;
                            }
                        }

                        // Combine Name plus any other textual candidates (AutomationId, HelpText, etc.)
                        var combinedText = loweredName;
                        if (element.TextCandidates != null && element.TextCandidates.Count > 0)
                        {
                            combinedText += " " + string.Join(" ", element.TextCandidates).ToLowerInvariant();
                        }

                        var allowFuzzy = AppSettings.Instance.VisualTargeting.AllowFuzzyMatching;
                        if (!MatchesTokens(combinedText, loweredPhrase, allowFuzzy))
                        {
                            Logger.LogDebug($"TryFromUiAutomation: skipping element '{name}' - not all tokens found. Tokens: {string.Join(", ", tokens)}, Element text: '{combinedText}'.");
                            continue;
                        }

                        score = 0.55;
                        if (string.Equals(combinedText, loweredPhrase, StringComparison.OrdinalIgnoreCase))
                        {
                            score = 0.95;
                        }
                        else if (combinedText.StartsWith(loweredPhrase, StringComparison.OrdinalIgnoreCase))
                        {
                            score = 0.86;
                        }
                        else if (ContainsWholePhrase(combinedText, loweredPhrase))
                        {
                            score = 0.78;
                        }
                    }

                    matches.Add(new VisualTargetCandidate
                    {
                        Label = name,
                        Bounds = element.Bounds,
                        Confidence = score,
                        Reason = $"UIA match '{name}'",
                        Source = "uia"
                    });
                    Logger.LogDebug($"TryFromUiAutomation: added match '{name}' with score {score:0.00}.");
                }

                Logger.LogDebug($"TryFromUiAutomation: completed, found {matches.Count} matches for '{phrase}'.");
                return matches;
            }
            catch (Exception ex)
            {
                Logger.LogError($"VisualTargetingService UIA fallback failed: {ex.Message}");
                return new List<VisualTargetCandidate>();
            }
        }

        private static List<VisualTargetCandidate> TryFromLocalOcr(string phrase, ScreenCaptureResult? preCapture = null)
        {
            try
            {
                // Run OCR on a background thread to avoid SynchronizationContext deadlock
                // The UI thread might be blocked waiting, and OCR operations need to marshal back to UI context
                List<VisualTargetCandidate> candidates;
                if (preCapture != null)
                {
                    candidates = Task.Run(async () => await LocalOcrService.FindCandidatesAsync(phrase, preCapture)).GetAwaiter().GetResult();
                }
                else
                {
                    candidates = Task.Run(async () => await LocalOcrService.FindCandidatesAsync(phrase)).GetAwaiter().GetResult();
                }

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

        private static bool ShouldPreferOcrForApp(string processName)
        {
            if (string.IsNullOrWhiteSpace(processName))
            {
                return false;
            }

            return string.Equals(processName, "steam", StringComparison.OrdinalIgnoreCase)
                || string.Equals(processName, "steamwebhelper", StringComparison.OrdinalIgnoreCase)
                || processName.Contains("steam", StringComparison.OrdinalIgnoreCase);
        }

        private static string ParseRankOnlyQuery(string phrase)
        {
            if (string.IsNullOrWhiteSpace(phrase))
            {
                return string.Empty;
            }

            string? detectedRank = null;
            var tokenSeparators = new[] { ' ', '\t', '\r', '\n' };
            var punctuation = new[] { ',', '.', ';', ':', '!', '?', '"', '\'' };

            foreach (var rawToken in phrase.Split(tokenSeparators, StringSplitOptions.RemoveEmptyEntries))
            {
                var token = rawToken.Trim().Trim(punctuation);
                if (token.Length == 0)
                {
                    continue;
                }

                if (string.Equals(token, "of", StringComparison.OrdinalIgnoreCase))
                {
                    break;
                }

                var normalized = NormalizeRank(token);
                if (string.IsNullOrWhiteSpace(normalized))
                {
                    continue;
                }

                if (detectedRank != null)
                {
                    return string.Empty;
                }

                detectedRank = normalized;
            }

            return detectedRank ?? string.Empty;
        }

        private static List<VisualTargetCandidate> BuildStackCardCandidates(
            System.Drawing.Rectangle stackBounds,
            List<CardValue> cardsInName,
            Func<CardValue, bool> predicate,
            double confidence,
            string reason)
        {
            var candidates = new List<VisualTargetCandidate>();
            if (cardsInName.Count == 0 || stackBounds.Width <= 0 || stackBounds.Height <= 0)
            {
                return candidates;
            }

            var estimatedCardBounds = EstimateStackCardBounds(stackBounds, cardsInName.Count);
            var count = Math.Min(cardsInName.Count, estimatedCardBounds.Count);
            for (var index = 0; index < count; index++)
            {
                var card = cardsInName[index];
                if (!predicate(card))
                {
                    continue;
                }

                candidates.Add(new VisualTargetCandidate
                {
                    Label = $"{card.Rank} of {card.Suit}",
                    Bounds = estimatedCardBounds[index],
                    Confidence = confidence,
                    Reason = reason,
                    Source = "uia"
                });
            }

            return candidates;
        }

        private static List<System.Drawing.Rectangle> EstimateStackCardBounds(System.Drawing.Rectangle stackBounds, int cardCount)
        {
            var results = new List<System.Drawing.Rectangle>();
            if (cardCount <= 0 || stackBounds.Width <= 0 || stackBounds.Height <= 0)
            {
                return results;
            }

            if (cardCount == 1)
            {
                results.Add(stackBounds);
                return results;
            }

            var estimatedCardHeight = (int)Math.Round(stackBounds.Width * 1.45);
            estimatedCardHeight = Math.Max(stackBounds.Width, Math.Min(stackBounds.Height, estimatedCardHeight));

            var overlapSpan = stackBounds.Height - estimatedCardHeight;
            var rowOffset = overlapSpan > 0
                ? overlapSpan / (double)(cardCount - 1)
                : stackBounds.Height / (double)cardCount;
            rowOffset = Math.Max(14, rowOffset);

            for (var index = 0; index < cardCount; index++)
            {
                var top = stackBounds.Top + (int)Math.Round(index * rowOffset);
                if (top >= stackBounds.Bottom)
                {
                    top = stackBounds.Bottom - 1;
                }

                var height = index == cardCount - 1
                    ? Math.Min(estimatedCardHeight, stackBounds.Bottom - top)
                    : Math.Min(Math.Max(24, (int)Math.Round(rowOffset) + 10), stackBounds.Bottom - top);
                height = Math.Max(1, height);

                results.Add(new System.Drawing.Rectangle(stackBounds.Left, top, stackBounds.Width, height));
            }

            return results;
        }

        private static bool ContainsWholePhrase(string text, string phrase)
        {
            if (string.IsNullOrWhiteSpace(text) || string.IsNullOrWhiteSpace(phrase))
            {
                return false;
            }

            var pattern = $@"\b{Regex.Escape(phrase)}\b";
            return Regex.IsMatch(text, pattern, RegexOptions.IgnoreCase);
        }

        // Public helper used by tests and matching logic: checks whether all tokens
        // from the phrase are present in the text. If allowFuzzy is true, small
        // edit-distance (Levenshtein) matches are accepted for individual tokens.
        public static bool MatchesTokens(string text, string phrase, bool allowFuzzy = false)
        {
            if (string.IsNullOrWhiteSpace(phrase) || string.IsNullOrWhiteSpace(text)) return false;

            var loweredText = text.ToLowerInvariant();
            var compactText = Regex.Replace(loweredText, "[^a-z0-9]+", string.Empty);
            var compactPhrase = Regex.Replace(phrase.ToLowerInvariant(), "[^a-z0-9]+", string.Empty);
            var words = Regex.Split(loweredText, "\\W+").Where(w => !string.IsNullOrWhiteSpace(w)).ToArray();
            var tokens = phrase.ToLowerInvariant()
                .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Distinct()
                .ToArray();

            // Handle compact/camel labels where separators are missing in UI text,
            // e.g. phrase "natural commands" vs label "NaturalCommands".
            if (tokens.Length >= 2
                && !string.IsNullOrWhiteSpace(compactPhrase)
                && compactText.Contains(compactPhrase, StringComparison.Ordinal))
            {
                return true;
            }

            foreach (var token in tokens)
            {
                if (string.IsNullOrWhiteSpace(token)) continue;

                // exact whole-word match
                if (Regex.IsMatch(loweredText, $@"\b{Regex.Escape(token)}\b", RegexOptions.IgnoreCase))
                {
                    continue;
                }

                // For single-word app targets, allow a close word-level match even when fuzzy is off.
                // This handles common cases like "llama"/"lama" targeting "Ollama".
                if (tokens.Length == 1 && token.Length >= 4)
                {
                    bool closeWordMatch = words.Any(w =>
                        (w.Contains(token, StringComparison.OrdinalIgnoreCase) && (w.Length - token.Length) <= 2)
                        || LevenshteinDistance(w, token) <= 1);

                    if (closeWordMatch)
                    {
                        continue;
                    }
                }

                if (!allowFuzzy)
                {
                    return false;
                }

                // fuzzy: compare token to individual words from the text
                bool matched = false;
                foreach (var w in words)
                {
                    var dist = LevenshteinDistance(w, token);
                    if (dist <= 1)
                    {
                        matched = true;
                        break;
                    }
                }

                if (!matched) return false;
            }

            return true;
        }

        private static int LevenshteinDistance(string a, string b)
        {
            if (a == b) return 0;
            if (string.IsNullOrEmpty(a)) return b?.Length ?? 0;
            if (string.IsNullOrEmpty(b)) return a.Length;

            var la = a.Length;
            var lb = b.Length;
            var d = new int[la + 1, lb + 1];
            for (int i = 0; i <= la; i++) d[i, 0] = i;
            for (int j = 0; j <= lb; j++) d[0, j] = j;

            for (int i = 1; i <= la; i++)
            {
                for (int j = 1; j <= lb; j++)
                {
                    int cost = (a[i - 1] == b[j - 1]) ? 0 : 1;
                    d[i, j] = Math.Min(Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1), d[i - 1, j - 1] + cost);
                }
            }

            return d[la, lb];
        }

        /// <summary>
        /// Apply a set of normalizations to the incoming phrase before attempting
        /// any form of visual targeting.  This is a public helper so unit tests can
        /// verify behavior.
        /// </summary>
        public static string NormalizePhrase(string phrase)
        {
            if (string.IsNullOrWhiteSpace(phrase))
            {
                return string.Empty;
            }

            var text = phrase.Trim();
            text = text.Trim('"', '\'', '.', ',', '!', '?', ';', ':');

            // fix common mis-hearings of the connector
            text = NormalizeCardConnectorMisrecognitions(text);

            // handle the very common speech-to-text error where "spades" comes
            // back as "space".  Only convert when it looks like a card phrase
            // to avoid mangling unrelated text.
            text = Regex.Replace(text, @"\b(of|off)\s+space\b", "of spades", RegexOptions.IgnoreCase);

            if (!string.Equals(text, phrase, StringComparison.OrdinalIgnoreCase))
            {
                Logger.LogDebug($"VisualTargetingService: normalized phrase from '{phrase}' to '{text}'.");
            }

            return text;
        }

        private static string NormalizeCardConnectorMisrecognitions(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return text;
            }

            var normalized = CardConnectorMisrecognitionRegex.Replace(text, "${rank} of ${suit}");
            if (!string.Equals(normalized, text, StringComparison.OrdinalIgnoreCase))
            {
                Logger.LogDebug($"VisualTargetingService: normalized card phrase from '{text}' to '{normalized}'.");
            }

            return normalized;
        }

        private static CardValue? ParseCardFromPhrase(string text)
        {
            var cards = ParseCardsFromText(text);
            return cards.Count == 1 ? cards[0] : null;
        }

        private static List<CardValue> ParseCardsFromText(string text)
        {
            var cards = new List<CardValue>();
            if (string.IsNullOrWhiteSpace(text))
            {
                return cards;
            }

            foreach (Match match in CardPhraseRegex.Matches(text))
            {
                var rank = NormalizeRank(match.Groups["rank"].Value);
                var suit = NormalizeSuit(match.Groups["suit"].Value);
                if (string.IsNullOrWhiteSpace(rank) || string.IsNullOrWhiteSpace(suit))
                {
                    continue;
                }

                cards.Add(new CardValue(rank, suit));
            }

            return cards;
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

        private sealed record CardValue(string Rank, string Suit);
    }
}
