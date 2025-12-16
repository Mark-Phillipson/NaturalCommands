using System;
using System.IO;
using Microsoft.Win32;
using Xunit;

namespace NaturalCommands_NET.Tests
{
    [Collection("SteamRegistry")]
    public class NaturalLanguageInterpreterSteamFallbackTests
    {
        [Fact]
        public async System.Threading.Tasks.Task PlayCommandResolvesToSteamUriWhenAIReturnsExe()
        {
            var temp = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            string? originalSteamPath = null;
            try
            {
                // create fake steam installation with Company of Heroes 2 manifest (appid 231430)
                var steamapps = Path.Combine(temp, "steamapps");
                Directory.CreateDirectory(steamapps);
                File.WriteAllText(Path.Combine(steamapps, "appmanifest_231430.acf"), "\"appid\"\t\"231430\"\n\"name\"\t\"Company of Heroes 2\"\n");

                // set registry SteamPath to our temp
                try
                {
                    originalSteamPath = (string?)Registry.GetValue("HKEY_CURRENT_USER\\Software\\Valve\\Steam", "SteamPath", null);
                    using var key = Registry.CurrentUser.CreateSubKey("Software\\Valve\\Steam");
                    key.SetValue("SteamPath", temp, RegistryValueKind.String);
                }
                catch { }

                var interpreter = new NaturalCommands.NaturalLanguageInterpreter();

                // Simulate the user saying the slightly misspelled phrase that previously hit AI fallback
                var action = await interpreter.InterpretAsync("play company of heroes to");
                Assert.NotNull(action);
                Assert.IsType<NaturalCommands.LaunchAppAction>(action);
                var launch = (NaturalCommands.LaunchAppAction)action;
                Assert.Contains("steam://rungameid/231430", launch.AppExe, StringComparison.OrdinalIgnoreCase);
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
