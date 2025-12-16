using NaturalCommands.Helpers;
using Xunit;

namespace NaturalCommands_NET.Tests
{
    public class VoiceDictationBehaviorTests
    {
        [Fact]
        public void DebounceGate_FiresOnlyAfterDelay()
        {
            var gate = new DebounceGate(900);

            gate.MarkChange(0);
            Assert.False(gate.TryConsume(899));
            Assert.True(gate.TryConsume(900));
            Assert.False(gate.TryConsume(901));
        }

        [Fact]
        public void DebounceGate_ResetsWhenMoreChangesArrive()
        {
            var gate = new DebounceGate(500);

            gate.MarkChange(0);
            gate.MarkChange(200);

            Assert.False(gate.TryConsume(699));
            Assert.True(gate.TryConsume(700));
        }

        [Fact]
        public void TemporaryMarqueeOverride_ActiveUntilExpiry()
        {
            var ov = new TemporaryMarqueeOverride();

            ov.Set("Listening", 1000, 0);
            Assert.True(ov.IsActive(999));
            Assert.False(ov.IsActive(1000));
        }

        [Fact]
        public void TemporaryMarqueeOverride_ReturnsMessageOnlyWhenActive()
        {
            var ov = new TemporaryMarqueeOverride();

            ov.Set("Hello", 10, 100);
            Assert.Equal("Hello", ov.GetMessage(109));
            Assert.Null(ov.GetMessage(110));
        }
    }
}
