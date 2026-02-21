using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using OpenAI.Chat;

namespace NaturalCommands.Helpers
{
    public static class TalonCommandMatcher
    {
        public static async Task<string?> MatchAsync(string utterance, IReadOnlyList<string> candidates)
        {
            if (string.IsNullOrWhiteSpace(utterance) || candidates == null || candidates.Count == 0)
            {
                return null;
            }

            var normalizedUtterance = Normalize(utterance);
            Logger.LogInfo($"[TalonMatcher] Matching utterance: '{utterance}' (normalized: '{normalizedUtterance}') against {candidates.Count} candidates");
            
            var exact = candidates.FirstOrDefault(c => string.Equals(Normalize(c), normalizedUtterance, StringComparison.Ordinal));
            if (!string.IsNullOrWhiteSpace(exact))
            {
                Logger.LogInfo($"[TalonMatcher] Exact match found: '{exact}'");
                return exact;
            }

            // Pre-compute token usage frequency for specificity scoring
            var tokenFrequency = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (var cmd in candidates)
            {
                var tokens = Normalize(cmd).Split(' ', StringSplitOptions.RemoveEmptyEntries);
                foreach (var token in tokens)
                {
                    if (!tokenFrequency.ContainsKey(token))
                        tokenFrequency[token] = 0;
                    tokenFrequency[token]++;
                }
            }

            var ranked = candidates
                .Select(c => 
                {
                    var normalizedC = Normalize(c);
                    var score = Score(normalizedUtterance, normalizedC, tokenFrequency);
                    var matchCount = CountMatchingTokens(normalizedUtterance, normalizedC);
                    return new RankedCandidate(c, score, matchCount);
                })
                .OrderByDescending(c => c.Score)
                .ThenByDescending(c => c.MatchCount)  // Tiebreaker: prefer more matching tokens
                .ThenBy(c => c.Command, StringComparer.OrdinalIgnoreCase)
                .Take(40)
                .ToList();

            if (ranked.Count == 0)
            {
                Logger.LogInfo($"[TalonMatcher] No candidates ranked");
                return null;
            }

            var top = ranked[0];
            Logger.LogInfo($"[TalonMatcher] Top match: '{top.Command}' with score {top.Score:0.000}. Top 5: {string.Join(", ", ranked.Take(5).Select(r => $"'{r.Command}'({r.Score:0.000})"))}")
;
            var talonSettings = Models.AppSettings.Instance.Talon;

            if (top.Score >= 0.95)
            {
                Logger.LogInfo($"[TalonMatcher] Score {top.Score:0.000} >= 0.95, returning confident fuzzy match");
                return top.Command;
            }

            // Try AI-assisted semantic matching (preferred over pure fuzzy scoring)
            var apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
            if (!string.IsNullOrWhiteSpace(apiKey))
            {
                try
                {
                    Logger.LogInfo($"[TalonMatcher] Attempting AI semantic matching for: '{utterance}'");
                    var modelName = Models.AppSettings.Instance.AI.ModelName;
                    var client = new ChatClient(modelName, apiKey);
                    var listForPrompt = string.Join("\n", ranked.Select((r, index) => $"{index + 1}. {r.Command}"));
                    var prompt =
                        "Choose the single best Talon command for this user utterance from the provided candidates. " +
                        "Prioritize semantic meaning: ignore grammar/ordering and find the closest meaning. " +
                        "Return strict JSON only with fields command (string) and confidence (0..1). " +
                        "If no candidate is suitable, set command to empty string and confidence to 0.";

                    var completion = await client.CompleteChatAsync(new List<ChatMessage>
                    {
                        new SystemChatMessage(prompt),
                        new UserChatMessage($"Utterance: {utterance}\nCandidates:\n{listForPrompt}")
                    });

                    var content = completion.Value.Content.FirstOrDefault()?.Text;
                    var parsed = ParseAiResponse(content);
                    if (parsed != null)
                    {
                        var chosen = ranked.FirstOrDefault(r => string.Equals(r.Command, parsed.Value.Command, StringComparison.OrdinalIgnoreCase));
                        if (!string.IsNullOrWhiteSpace(chosen.Command) && parsed.Value.Confidence >= talonSettings.MinAIConfidence)
                        {
                            Logger.LogInfo($"[TalonMatcher] AI matched '{chosen.Command}' with confidence {parsed.Value.Confidence:0.00}");
                            return chosen.Command;
                        }
                    }
                    else
                    {
                        Logger.LogInfo($"[TalonMatcher] AI response parse failed");
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogWarning($"[TalonMatcher] AI semantic matching failed: {ex.Message}");
                }
            }
            else
            {
                Logger.LogInfo($"[TalonMatcher] No API key available for semantic matching");
            }

            // Fallback: use fuzzy score threshold if AI is unavailable or failed
            var result = top.Score >= talonSettings.MinFallbackMatchScore ? top.Command : null;
            Logger.LogInfo($"[TalonMatcher] Using fuzzy fallback. Top score {top.Score:0.000} vs threshold {talonSettings.MinFallbackMatchScore}. Result: {(result == null ? "null" : $"'{result}'")}");
            return result;
        }

        public static string Normalize(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var normalized = value.ToLowerInvariant();
            normalized = Regex.Replace(normalized, "[^a-z0-9\\s]", " ");
            normalized = Regex.Replace(normalized, "\\s+", " ").Trim();
            return normalized;
        }

        private static double Score(string a, string b, Dictionary<string, int> tokenFrequency)
        {
            if (string.IsNullOrWhiteSpace(a) || string.IsNullOrWhiteSpace(b)) return 0;
            if (string.Equals(a, b, StringComparison.Ordinal)) return 1;

            var aTokens = a.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            var bTokens = b.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            if (aTokens.Length == 0 || bTokens.Length == 0) return 0;

            var aSet = new HashSet<string>(aTokens, StringComparer.Ordinal);
            var bSet = new HashSet<string>(bTokens, StringComparer.Ordinal);
            
            // Count exact token matches
            var overlap = aSet.Count(t => bSet.Contains(t));
            
            // Base scoring: recall (coverage of candidate in utterance) + precision (coverage of utterance in candidate)
            var tokenRecall = overlap / (double)bSet.Count;  // How much of candidate is in utterance
            var tokenPrecision = overlap / (double)aSet.Count; // How much of utterance matches candidate
            
            // Weight recall more heavily - if all candidate tokens are in utterance, it's likely a match
            // even if utterance has extra words
            var baseScore = (tokenRecall * 0.7 + tokenPrecision * 0.3);

            // Bonus: if candidate is short and all tokens match, award high score
            if (bTokens.Length <= 2 && tokenRecall == 1.0)
            {
                return 0.9; // Near-perfect for all candidate tokens matching
            }
            
            // Bonus: if candidate is short and most of it matches, award higher score
            if (bTokens.Length <= 3 && tokenRecall >= 0.5)
            {
                baseScore = Math.Max(baseScore, 0.7 + (tokenRecall * 0.25));
            }

            var containsBoost = b.Contains(a, StringComparison.Ordinal) || a.Contains(b, StringComparison.Ordinal) ? 0.1 : 0;
            return Math.Min(1.0, baseScore + containsBoost);
        }

        private static (string Command, double Confidence)? ParseAiResponse(string? content)
        {
            if (string.IsNullOrWhiteSpace(content))
            {
                return null;
            }

            try
            {
                var trimmed = content.Trim();
                if (trimmed.StartsWith("```", StringComparison.Ordinal))
                {
                    var start = trimmed.IndexOf('{');
                    var end = trimmed.LastIndexOf('}');
                    if (start >= 0 && end > start)
                    {
                        trimmed = trimmed.Substring(start, end - start + 1);
                    }
                }

                using var doc = JsonDocument.Parse(trimmed);
                var root = doc.RootElement;
                var command = root.TryGetProperty("command", out var c) ? c.GetString() ?? string.Empty : string.Empty;
                var confidence = root.TryGetProperty("confidence", out var conf) && conf.ValueKind == JsonValueKind.Number
                    ? conf.GetDouble()
                    : 0;
                return (command, confidence);
            }
            catch
            {
                return null;
            }
        }

        private readonly record struct RankedCandidate(string Command, double Score, int MatchCount);
        
        private static int CountMatchingTokens(string a, string b)
        {
            if (string.IsNullOrWhiteSpace(a) || string.IsNullOrWhiteSpace(b)) return 0;
            
            var aTokens = a.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            var bTokens = b.Split(' ', StringSplitOptions.RemoveEmptyEntries);
            
            var aSet = new HashSet<string>(aTokens, StringComparer.Ordinal);
            var bSet = new HashSet<string>(bTokens, StringComparer.Ordinal);
            
            return aSet.Count(t => bSet.Contains(t));
        }
    }
}
