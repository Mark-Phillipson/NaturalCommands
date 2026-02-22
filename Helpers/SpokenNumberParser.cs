using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace NaturalCommands.Helpers
{
    internal static class SpokenNumberParser
    {
        private static readonly Dictionary<string, int> NumberWords = new(StringComparer.OrdinalIgnoreCase)
        {
            { "one", 1 }, { "won", 1 },
            { "two", 2 }, { "to", 2 }, { "too", 2 },
            { "three", 3 },
            { "four", 4 }, { "for", 4 },
            { "five", 5 },
            { "six", 6 },
            { "seven", 7 },
            { "eight", 8 }, { "ate", 8 },
            { "nine", 9 },
            { "ten", 10 },
            { "eleven", 11 },
            { "twelve", 12 }
        };

        private static readonly Dictionary<string, int> OrdinalWords = new(StringComparer.OrdinalIgnoreCase)
        {
            { "first", 1 },
            { "second", 2 },
            { "third", 3 },
            { "fourth", 4 },
            { "fifth", 5 },
            { "sixth", 6 },
            { "seventh", 7 },
            { "eighth", 8 },
            { "ninth", 9 },
            { "tenth", 10 },
            { "eleventh", 11 },
            { "twelfth", 12 }
        };

        public static bool TryParsePositive(string? value, out int number)
        {
            number = 0;
            if (string.IsNullOrWhiteSpace(value))
                return false;

            var token = value.Trim();

            if (int.TryParse(token, out var parsed) && parsed > 0)
            {
                number = parsed;
                return true;
            }

            var ordinalMatch = Regex.Match(token, @"^(\d+)(st|nd|rd|th)$", RegexOptions.IgnoreCase);
            if (ordinalMatch.Success && int.TryParse(ordinalMatch.Groups[1].Value, out parsed) && parsed > 0)
            {
                number = parsed;
                return true;
            }

            if (NumberWords.TryGetValue(token, out parsed) && parsed > 0)
            {
                number = parsed;
                return true;
            }

            if (OrdinalWords.TryGetValue(token, out parsed) && parsed > 0)
            {
                number = parsed;
                return true;
            }

            return false;
        }
    }
}
