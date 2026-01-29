using System.Linq;
using Xunit;

namespace NaturalCommands_NET.Tests
{
    public class NaturalLanguageInterpreterAutoClickCommandsTests
    {
        [Fact]
        public void AvailableCommands_Contains_AutoClickSpeedCommands()
        {
            var cmds = NaturalCommands.NaturalLanguageInterpreter.AvailableCommands.Select(c => c.Command).ToList();
            Assert.Contains("auto click faster", cmds);
            Assert.Contains("auto click slower", cmds);
            Assert.Contains("auto click speed up", cmds);
            Assert.Contains("auto click slow down", cmds);
        }
    }
}