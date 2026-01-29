using System.Linq;
using Xunit;

namespace NaturalCommands_NET.Tests
{
    public class NaturalLanguageInterpreterAvailableCommandsTests
    {
        [Fact]
        public void AvailableCommands_Contains_OpenSettings()
        {
            var found = NaturalCommands.NaturalLanguageInterpreter.AvailableCommands.Any(c => c.Command == "open settings");
            Assert.True(found, "Expected AvailableCommands to contain 'open settings'.");
        }
    }
}