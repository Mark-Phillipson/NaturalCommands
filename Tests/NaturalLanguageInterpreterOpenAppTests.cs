using Xunit;

namespace NaturalCommands_NET.Tests
{
    public class NaturalLanguageInterpreterOpenAppTests
    {
        [Fact]
        public async System.Threading.Tasks.Task OpenSteamApplication_MapsToSteamUri()
        {
            var interpreter = new NaturalCommands.NaturalLanguageInterpreter();

            var action = await interpreter.InterpretAsync("please open the steam application");

            Assert.NotNull(action);
            var launch = Assert.IsType<NaturalCommands.LaunchAppAction>(action);
            Assert.Equal("steam://open/main", launch.AppExe);
        }

        [Fact]
        public async System.Threading.Tasks.Task OpenSteamWithConversationalSuffix_MapsToSteamUri()
        {
            var interpreter = new NaturalCommands.NaturalLanguageInterpreter();

            var action = await interpreter.InterpretAsync("please open steam up for me");

            Assert.NotNull(action);
            var launch = Assert.IsType<NaturalCommands.LaunchAppAction>(action);
            Assert.Equal("steam://open/main", launch.AppExe);
        }

        [Theory]
        [InlineData("please open calculator up for me", "calc.exe")]
        [InlineData("please open chrome for me", "chrome.exe")]
        [InlineData("please open visual studio up for me", "devenv.exe")]
        [InlineData("please open microsoft paint for me", "mspaint.exe")]
        [InlineData("please open the terminal application for me", "wt.exe")]
        public async System.Threading.Tasks.Task OpenMappedApplications_WithConversationalPhrases_MapCorrectly(string phrase, string expectedExecutable)
        {
            var interpreter = new NaturalCommands.NaturalLanguageInterpreter();

            var action = await interpreter.InterpretAsync(phrase);

            Assert.NotNull(action);
            var launch = Assert.IsType<NaturalCommands.LaunchAppAction>(action);
            Assert.Equal(expectedExecutable, launch.AppExe);
        }

        [Theory]
        [InlineData("please open up calculator", "calc.exe")]
        [InlineData("please open up chrome", "chrome.exe")]
        [InlineData("please open up visual studio", "devenv.exe")]
        [InlineData("please open up microsoft paint", "mspaint.exe")]
        [InlineData("please open up terminal", "wt.exe")]
        public async System.Threading.Tasks.Task OpenMappedApplications_WithOpenUpPattern_MapCorrectly(string phrase, string expectedExecutable)
        {
            var interpreter = new NaturalCommands.NaturalLanguageInterpreter();

            var action = await interpreter.InterpretAsync(phrase);

            Assert.NotNull(action);
            var launch = Assert.IsType<NaturalCommands.LaunchAppAction>(action);
            Assert.Equal(expectedExecutable, launch.AppExe);
        }

        [Theory]
        [InlineData("open calculator now please", "calc.exe")]
        [InlineData("open chrome now please", "chrome.exe")]
        [InlineData("open visual studio now please", "devenv.exe")]
        [InlineData("open microsoft paint now please", "mspaint.exe")]
        [InlineData("open terminal now please", "wt.exe")]
        [InlineData("open steam now please", "steam://open/main")]
        public async System.Threading.Tasks.Task OpenMappedApplications_WithNowPleaseSuffix_MapCorrectly(string phrase, string expectedExecutable)
        {
            var interpreter = new NaturalCommands.NaturalLanguageInterpreter();

            var action = await interpreter.InterpretAsync(phrase);

            Assert.NotNull(action);
            var launch = Assert.IsType<NaturalCommands.LaunchAppAction>(action);
            Assert.Equal(expectedExecutable, launch.AppExe);
        }

        [Fact]
        public async System.Threading.Tasks.Task OpenUnknownApplication_UsesFocusFallback()
        {
            var interpreter = new NaturalCommands.NaturalLanguageInterpreter();

            var action = await interpreter.InterpretAsync("please open frobnicator now please");

            Assert.NotNull(action);
            var launch = Assert.IsType<NaturalCommands.LaunchAppAction>(action);
            Assert.Equal("focus-fallback", launch.AppExe);
        }
    }
}
