using System.Threading.Tasks;

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

            return new NaturalCommands.RunTalonCommandAction(text, match, source);
        }
    }
}
