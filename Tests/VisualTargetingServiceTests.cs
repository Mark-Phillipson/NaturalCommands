using Xunit;
using NaturalCommands.Helpers;
using NaturalCommands.Models;

namespace NaturalCommands_NET.Tests
{
    public class VisualTargetingServiceTests
    {
        [Theory]
        [InlineData("king of clubs", new[] {"king", "of", "clubs", "k", "♣"})]
        [InlineData("Ace of Spades", new[] {"ace", "of", "spades", "a", "♠"})]
        [InlineData("10 of hearts", new[] {"10", "of", "hearts", "10", "♥"})]
        public void TokenList_IncludesCardSynonyms(string phrase, string[] expectedTokens)
        {
            var tokens = LocalOcrService.BuildTokenListForPhrase(phrase);
            foreach (var expected in expectedTokens)
            {
                Assert.Contains(expected, tokens);
            }
        }

        [Theory]
        [InlineData("queen of space", "queen of spades")]
        [InlineData("Queen off spades", "Queen of spades")]
        [InlineData("king of spades", "king of spades")]
        public void NormalizePhrase_FixesCommonMishearings(string input, string expected)
        {
            var normalized = VisualTargetingService.NormalizePhrase(input);
            Assert.Equal(expected, normalized, ignoreCase: true);
        }

        [Fact]
        public void IdentifyCandidates_SpaceAndSpadesAreEquivalent()
        {
            var r1 = VisualTargetingService.IdentifyCandidates("queen of space");
            var r2 = VisualTargetingService.IdentifyCandidates("queen of spades");
            Assert.Equal(r2.Count, r1.Count);
        }

        [Fact]
        public void CloudVisionCandidates_ReturnsEmpty_WhenUnavailable()
        {
            // simulate cloud vision disabled
            System.Environment.SetEnvironmentVariable("OPENAI_API_KEY", null);
            AppSettings.Instance.VisualTargeting.MaxCloudCallsPerDay = 0;
            var results = VisualTargetingService.GetCloudVisionCandidates("any phrase");
            Assert.Empty(results);
        }

        [Fact]
        public void MatchesTokens_BasicAndFuzzy()
        {
            // exact tokens present
            Assert.True(VisualTargetingService.MatchesTokens("File Save As", "save as", false));
            Assert.False(VisualTargetingService.MatchesTokens("File Save As", "open file", false));

            // small typo allowed when fuzzy matching is enabled
            Assert.True(VisualTargetingService.MatchesTokens("File Save As", "sve", true));
        }
    }
}
