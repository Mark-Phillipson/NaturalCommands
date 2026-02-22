using System;
using System.Drawing;
using System.Text.RegularExpressions;
using NaturalCommands.Models;

namespace NaturalCommands.Helpers
{
    internal static class VisualTargetClickPointResolver
    {
        private static readonly Regex CardPhraseRegex = new(@"\b(?<rank>ace|king|queen|jack|10|[2-9]|two|three|four|five|six|seven|eight|nine|ten)\s+(of\s+)?(?<suit>spades?|hearts?|diamonds?|clubs?)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static Point Resolve(VisualTargetCandidate candidate, string query)
        {
            if (candidate == null || candidate.Bounds.Width <= 0 || candidate.Bounds.Height <= 0)
            {
                return Point.Empty;
            }

            if (!LooksLikeCardQuery(query) && !LooksLikeCardQuery(candidate.Label))
            {
                return candidate.Center;
            }

            var topOffset = Math.Max(8, Math.Min(26, candidate.Bounds.Height / 6));
            var y = Math.Min(candidate.Bounds.Bottom - 1, candidate.Bounds.Top + topOffset);
            var x = candidate.Bounds.Left + (candidate.Bounds.Width / 2);
            return new Point(x, y);
        }

        private static bool LooksLikeCardQuery(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            return CardPhraseRegex.IsMatch(value.Trim());
        }
    }
}