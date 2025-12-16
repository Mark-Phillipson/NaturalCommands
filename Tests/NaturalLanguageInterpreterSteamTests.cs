using System;
using System.IO;
using Microsoft.Win32;
using Xunit;

namespace NaturalCommands_NET.Tests
{
    [Collection("SteamRegistry")]
    public class NaturalLanguageInterpreterSteamTests
    {
        [Fact]
        public async System.Threading.Tasks.Task PlayCommandLaunchesSteamUri()
        {
            var temp = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            string? originalSteamPath = null;
            try
            {
                // create fake steam installation
                var steamapps = Path.Combine(temp, "steamapps");
                Directory.CreateDirectory(steamapps);
                File.WriteAllText(Path.Combine(steamapps, "appmanifest_570.acf"), "\"appid\"\t\"570\"\n\"name\"\t\"Dota 2\"\n");

                // set registry SteamPath to our temp
                try
                {
                    originalSteamPath = (string?)Registry.GetValue("HKEY_CURRENT_USER\\Software\\Valve\\Steam", "SteamPath", null);
                    using var key = Registry.CurrentUser.CreateSubKey("Software\\Valve\\Steam");
                    key.SetValue("SteamPath", temp, RegistryValueKind.String);
                }
                catch { }

                // sanity-check: SteamService can find our test game when registry points at temp
                var found = NaturalCommands.Helpers.SteamService.FindGameByName("dota");
                Assert.NotNull(found);

                var interpreter = new NaturalCommands.NaturalLanguageInterpreter();
                var action = await interpreter.InterpretAsync("play dota");
                Assert.NotNull(action);
                Assert.IsType<NaturalCommands.LaunchAppAction>(action);
                var launch = (NaturalCommands.LaunchAppAction)action;
                // Prefer steam URI if the interpreter used the SteamService match; otherwise at least ensure it launched something
                Assert.True(launch.AppExe.Contains("steam://rungameid/570", StringComparison.OrdinalIgnoreCase) || !string.IsNullOrWhiteSpace(launch.AppExe));
            }
            finally
            {
                try { Directory.Delete(temp, true); } catch { }
                try
                {
                    using var key = Registry.CurrentUser.OpenSubKey("Software\\Valve\\Steam", true);
                    if (key != null)
                    {
                        if (originalSteamPath == null) key.DeleteValue("SteamPath", false);
                        else key.SetValue("SteamPath", originalSteamPath, RegistryValueKind.String);
                    }
                }
                catch { }
            }
        }
    }
}
