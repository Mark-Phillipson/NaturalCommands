using NaturalCommands.Helpers;
using Xunit;

namespace NaturalCommands_NET.Tests
{
    public class TalonCommandRouterTests
    {
        [Fact]
        public void ResolveDispatchCommand_ReplacesSinglePlaceholderWithSpokenValue()
        {
            var resolved = TalonCommandRouter.ResolveDispatchCommand("please launch steam", "launch <user.launch_applications>");

            Assert.Equal("launch steam", resolved);
        }

        [Fact]
        public void ResolveDispatchCommand_LeavesCommandUnchangedWhenNoPlaceholder()
        {
            var resolved = TalonCommandRouter.ResolveDispatchCommand("launch steam", "launch steam");

            Assert.Equal("launch steam", resolved);
        }
    }
}
