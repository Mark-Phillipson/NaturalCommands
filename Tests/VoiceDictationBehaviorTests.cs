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

        [Fact]
        public void TemporaryMarqueeOverride_Clear_RemovesMessage()
        {
            var ov = new TemporaryMarqueeOverride();

            ov.Set("Hello", 1000, 0);
            Assert.Equal("Hello", ov.GetMessage(10));
            ov.Clear();
            Assert.Null(ov.GetMessage(10));
        }

        [Fact]
        public void IsCommandAvailable_DetectsCmd()
        {
            // cmd.exe should always be available on Windows
            var form = new DictationBoxMSP.VoiceDictationForm();
            var method = typeof(DictationBoxMSP.VoiceDictationForm)
                .GetMethod("IsCommandAvailable", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            Assert.NotNull(method);
            bool result = (bool)method.Invoke(form, new object[] { "cmd.exe" })!;
            Assert.True(result, "cmd.exe should be detected as available");
        }

        [Fact]
        public void LaunchTerminalWithCommand_DoesNotThrow()
        {
            var form = new DictationBoxMSP.VoiceDictationForm();
            var method = typeof(DictationBoxMSP.VoiceDictationForm)
                .GetMethod("LaunchTerminalWithCommand", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            Assert.NotNull(method);
            // using a harmless echo command; actual terminal may appear but test will not wait
            method.Invoke(form, new object[] { "echo hello" });
        }

        [Fact]
        public void BtnSendToCopilot_CopiesClipboard()
        {
            var th = new System.Threading.Thread(() =>
            {
                var form = new DictationBoxMSP.VoiceDictationForm();
                // set input text via reflection
                var txtField = typeof(DictationBoxMSP.VoiceDictationForm)
                    .GetField("txtInput", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(txtField);
                var txtBox = (System.Windows.Forms.TextBox)txtField.GetValue(form)!;
                txtBox.Text = "clipboard test";

                var method = typeof(DictationBoxMSP.VoiceDictationForm)
                    .GetMethod("BtnSendToCopilot_Click", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                Assert.NotNull(method);

                method.Invoke(form, new object[] { null, System.EventArgs.Empty });

                Assert.Equal("clipboard test", System.Windows.Forms.Clipboard.GetText());
            });
            th.SetApartmentState(System.Threading.ApartmentState.STA);
            th.Start();
            th.Join();
        }

        [Fact]
        public void GetCopilotCommand_AlwaysOpensCli()
        {
            var form = new DictationBoxMSP.VoiceDictationForm();
            var method = typeof(DictationBoxMSP.VoiceDictationForm)
                .GetMethod("GetCopilotCommand", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            Assert.NotNull(method);

            var cmd1 = (string)method.Invoke(form, new object[] { "hi", true, false })!;
            Assert.Equal("copilot", cmd1);

            var cmd2 = (string)method.Invoke(form, new object[] { "any text", false, true })!;
            Assert.Equal("gh copilot chat", cmd2);

            var cmd3 = (string)method.Invoke(form, new object[] { "multi\nline", true, false })!;
            Assert.Equal("copilot", cmd3);
        }
    }
}
