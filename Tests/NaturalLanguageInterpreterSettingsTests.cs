using System;
using Xunit;

namespace NaturalCommands_NET.Tests
{
    public class NaturalLanguageInterpreterSettingsTests
    {
        [Fact]
        public async System.Threading.Tasks.Task OpenSettingsMatches_OpenSettings()
        {
            var interpreter = new NaturalCommands.NaturalLanguageInterpreter();
            var action = await interpreter.InterpretAsync("open settings");
            Assert.NotNull(action);
            Assert.IsType<NaturalCommands.OpenSettingsAction>(action);
        }

        [Fact]
        public async System.Threading.Tasks.Task OpenSettingsMatches_ShowSettings()
        {
            var interpreter = new NaturalCommands.NaturalLanguageInterpreter();
            var action = await interpreter.InterpretAsync("show settings");
            Assert.NotNull(action);
            Assert.IsType<NaturalCommands.OpenSettingsAction>(action);
        }

        [Fact]
        public async System.Threading.Tasks.Task OpenSettingsMatches_SettingsOnly()
        {
            var interpreter = new NaturalCommands.NaturalLanguageInterpreter();
            var action = await interpreter.InterpretAsync("settings");
            Assert.NotNull(action);
            Assert.IsType<NaturalCommands.OpenSettingsAction>(action);
        }
    }
}