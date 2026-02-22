using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace NaturalCommands.Helpers
{
    public static class TalonCommandRouter
    {
        public static async Task<NaturalCommands.RunTalonCommandAction?> ResolveActionAsync(string text)
        {
            var settings = Models.AppSettings.Instance.Talon;
            if (!settings.EnableFallback)
            {
                return null;
            }

            var commands = TalonCommandCatalog.GetCommands();
            if (commands == null || commands.Count == 0)
            {
                Logger.LogWarning("TalonCommandRouter: Talon catalog is empty.");
                return null;
            }

            var contextPreferences = TalonCommandCatalog.GetContextPreferences();
            var match = await TalonCommandMatcher.MatchAsync(text, commands, contextPreferences);
            if (string.IsNullOrWhiteSpace(match))
            {
                return null;
            }

            var source = TalonCommandMatcher.Normalize(text) == TalonCommandMatcher.Normalize(match)
                ? "exact"
                : "ai-assisted";

            var dispatchCommand = ResolveDispatchCommand(text, match);
            return new NaturalCommands.RunTalonCommandAction(text, dispatchCommand, source);
        }

        public static string ResolveDispatchCommand(string utterance, string matchedCommand)
        {
            if (string.IsNullOrWhiteSpace(matchedCommand) || !matchedCommand.Contains('<'))
            {
                return matchedCommand;
            }

            var templateTokens = matchedCommand
                .Split(' ', System.StringSplitOptions.RemoveEmptyEntries)
                .ToList();
            if (templateTokens.Count == 0)
            {
                return matchedCommand;
            }

            var utteranceTokens = TokenizeUtterance(utterance);
            if (utteranceTokens.Count == 0)
            {
                return matchedCommand;
            }

            var outputTokens = new List<string>();
            var utteranceIndex = 0;

            for (var templateIndex = 0; templateIndex < templateTokens.Count; templateIndex++)
            {
                var token = templateTokens[templateIndex];
                if (IsPlaceholder(token))
                {
                    var nextLiteral = FindNextLiteralToken(templateTokens, templateIndex + 1);
                    var captureEnd = nextLiteral == null
                        ? utteranceTokens.Count
                        : FindLiteralIndex(utteranceTokens, nextLiteral, utteranceIndex);

                    if (captureEnd < utteranceIndex)
                    {
                        return matchedCommand;
                    }

                    var captured = utteranceTokens.Skip(utteranceIndex).Take(captureEnd - utteranceIndex).ToArray();
                    if (captured.Length == 0)
                    {
                        return matchedCommand;
                    }

                    outputTokens.Add(string.Join(" ", captured));
                    utteranceIndex = captureEnd;
                    continue;
                }

                var literalIndex = FindLiteralIndex(utteranceTokens, token, utteranceIndex);
                if (literalIndex < utteranceIndex)
                {
                    return matchedCommand;
                }

                outputTokens.Add(token);
                utteranceIndex = literalIndex + 1;
            }

            return outputTokens.Count == 0 ? matchedCommand : string.Join(" ", outputTokens);
        }

        private static List<string> TokenizeUtterance(string utterance)
        {
            if (string.IsNullOrWhiteSpace(utterance))
            {
                return new List<string>();
            }

            var normalized = utterance.ToLowerInvariant();
            normalized = Regex.Replace(normalized, "[^a-z0-9\\s]", " ");
            normalized = Regex.Replace(normalized, "\\s+", " ").Trim();

            var tokens = normalized.Split(' ', System.StringSplitOptions.RemoveEmptyEntries).ToList();

            while (tokens.Count > 0 && IsPolitePrefix(tokens[0]))
            {
                tokens.RemoveAt(0);
            }

            return tokens;
        }

        private static bool IsPolitePrefix(string token)
        {
            return token == "please"
                || token == "can"
                || token == "could"
                || token == "would"
                || token == "will"
                || token == "you"
                || token == "hey"
                || token == "hi"
                || token == "hello";
        }

        private static bool IsPlaceholder(string token)
        {
            return token.StartsWith("<", System.StringComparison.Ordinal)
                && token.EndsWith(">", System.StringComparison.Ordinal)
                && token.Length > 2;
        }

        private static string? FindNextLiteralToken(List<string> templateTokens, int startIndex)
        {
            for (var i = startIndex; i < templateTokens.Count; i++)
            {
                if (!IsPlaceholder(templateTokens[i]))
                {
                    return templateTokens[i];
                }
            }

            return null;
        }

        private static int FindLiteralIndex(List<string> utteranceTokens, string literalToken, int startIndex)
        {
            var normalizedLiteral = NormalizeToken(literalToken);
            if (string.IsNullOrWhiteSpace(normalizedLiteral))
            {
                return -1;
            }

            for (var i = startIndex; i < utteranceTokens.Count; i++)
            {
                if (utteranceTokens[i] == normalizedLiteral)
                {
                    return i;
                }
            }

            return -1;
        }

        private static string NormalizeToken(string token)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return string.Empty;
            }

            return Regex.Replace(token.ToLowerInvariant(), "[^a-z0-9]", string.Empty);
        }
    }
}
