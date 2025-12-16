using System;
using System.IO;
using System.Linq;
using Xunit;
using NaturalCommands.Helpers;

namespace NaturalCommands_NET.Tests
{
    public class SteamServiceTests
    {
        [Fact]
        public void ParsesAppmanifestFilesAndFindsGameByName()
        {
            var temp = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            try
            {
                var steamapps = Path.Combine(temp, "steamapps");
                Directory.CreateDirectory(steamapps);

                var acf = @"""
"""; // empty placeholder in case of formatting

                var manifest1 = "appmanifest_570.acf";
                File.WriteAllText(Path.Combine(steamapps, manifest1), "\"appid\"\t\"570\"\n\"name\"\t\"Dota 2\"\n");

                var manifest2 = "appmanifest_440.acf";
                File.WriteAllText(Path.Combine(steamapps, manifest2), "\"appid\"\t\"440\"\n\"name\"\t\"Team Fortress 2\"\n");

                var games = SteamService.GetInstalledGames(temp).ToList();
                Assert.Contains(games, g => g.AppId == "570" && g.Name == "Dota 2");
                Assert.Contains(games, g => g.AppId == "440" && g.Name == "Team Fortress 2");

                var found = SteamService.FindGameByName("dota", temp);
                Assert.NotNull(found);
                Assert.Equal("570", found!.AppId);
            }
            finally
            {
                try { Directory.Delete(temp, true); } catch { }
            }
        }

        [Fact]
        public void ReadsModernLibraryfoldersVdfAndScansAdditionalLibraries()
        {
            var temp = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            try
            {
                var steamRoot = Path.Combine(temp, "Steam");
                var defaultSteamapps = Path.Combine(steamRoot, "steamapps");
                Directory.CreateDirectory(defaultSteamapps);

                var altLibrary = Path.Combine(temp, "SteamLibrary");
                var altSteamapps = Path.Combine(altLibrary, "steamapps");
                Directory.CreateDirectory(altSteamapps);

                // Put a game manifest only in the alternate library
                File.WriteAllText(Path.Combine(altSteamapps, "appmanifest_231430.acf"), "\"appid\"\t\"231430\"\n\"name\"\t\"Company of Heroes 2\"\n");

                // Modern VDF format includes a nested block with a "path" key
                var vdf = "\"libraryfolders\"\n{\n  \"1\"\n  {\n    \"path\"\t\"" + altLibrary.Replace("\\", "\\\\") + "\"\n  }\n}\n";
                File.WriteAllText(Path.Combine(defaultSteamapps, "libraryfolders.vdf"), vdf);

                var found = SteamService.FindGameByName("company of heroes 2", steamRoot);
                Assert.NotNull(found);
                Assert.Equal("231430", found!.AppId);
            }
            finally
            {
                try { Directory.Delete(temp, true); } catch { }
            }
        }
    }
}
