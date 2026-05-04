using System.Collections.Generic;
using NaturalCommands;
using Xunit;

namespace NaturalCommands_NET.Tests
{
    public class WindowsNotificationListenerServiceTests
    {
        // ----------------------------------------------------------------
        // ParseCategoryRulesJson
        // ----------------------------------------------------------------

        [Fact]
        public void ParseCategoryRulesJson_MapsKnownApp()
        {
            var json = """[{"app":"Outlook","category":"warning"},{"default":"info"}]""";
            var map = WindowsNotificationListenerService.ParseCategoryRulesJson(json);

            Assert.Equal(TickerCategory.Warning, map["Outlook"]);
        }

        [Fact]
        public void ParseCategoryRulesJson_CaseInsensitiveAppLookup()
        {
            var json = """[{"app":"teams","category":"critical"}]""";
            var map = WindowsNotificationListenerService.ParseCategoryRulesJson(json);

            Assert.Equal(TickerCategory.Critical, map["Teams"]);
        }

        [Fact]
        public void ParseCategoryRulesJson_DefaultFallsBackToInfo_WhenMissing()
        {
            var json = """[{"app":"Chrome","category":"success"}]""";
            var map = WindowsNotificationListenerService.ParseCategoryRulesJson(json);

            Assert.Equal(TickerCategory.Info, map["default"]);
        }

        [Fact]
        public void ParseCategoryRulesJson_ParsesDefaultEntry()
        {
            var json = """[{"default":"warning"}]""";
            var map = WindowsNotificationListenerService.ParseCategoryRulesJson(json);

            Assert.Equal(TickerCategory.Warning, map["default"]);
        }

        [Fact]
        public void ParseCategoryRulesJson_IgnoresUnknownCategory()
        {
            var json = """[{"app":"SomeApp","category":"unknowncat"}]""";
            var map = WindowsNotificationListenerService.ParseCategoryRulesJson(json);

            Assert.False(map.ContainsKey("SomeApp"));
        }

        [Fact]
        public void ParseCategoryRulesJson_ReturnsInfoDefault_ForNullJson()
        {
            var map = WindowsNotificationListenerService.ParseCategoryRulesJson("null");

            Assert.Equal(TickerCategory.Info, map["default"]);
        }

        // ----------------------------------------------------------------
        // ParseIgnoreRulesJson
        // ----------------------------------------------------------------

        [Fact]
        public void ParseIgnoreRulesJson_ContainsListedApp()
        {
            var json = """[{"app":"Windows Security"},{"app":"UAC"}]""";
            var set = WindowsNotificationListenerService.ParseIgnoreRulesJson(json);

            Assert.Contains("Windows Security", set);
            Assert.Contains("UAC", set);
        }

        [Fact]
        public void ParseIgnoreRulesJson_IsCaseInsensitive()
        {
            var json = """[{"app":"windows security"}]""";
            var set = WindowsNotificationListenerService.ParseIgnoreRulesJson(json);

            Assert.Contains("Windows Security", set);
        }

        [Fact]
        public void ParseIgnoreRulesJson_SkipsEntriesWithoutAppKey()
        {
            var json = """[{"other":"value"},{"app":"Chrome"}]""";
            var set = WindowsNotificationListenerService.ParseIgnoreRulesJson(json);

            Assert.Single(set);
            Assert.Contains("Chrome", set);
        }

        [Fact]
        public void ParseIgnoreRulesJson_ReturnsEmpty_ForNullJson()
        {
            var set = WindowsNotificationListenerService.ParseIgnoreRulesJson("null");

            Assert.Empty(set);
        }

        // ----------------------------------------------------------------
        // MapAppToCategory (via internal test constructor)
        // ----------------------------------------------------------------

        [Fact]
        public void MapAppToCategory_ReturnsMatchedCategory()
        {
            var map = new Dictionary<string, TickerCategory>(System.StringComparer.OrdinalIgnoreCase)
            {
                ["Outlook"] = TickerCategory.Warning,
                ["default"] = TickerCategory.Info
            };
            var svc = new WindowsNotificationListenerService(map, new HashSet<string>());

            Assert.Equal(TickerCategory.Warning, svc.MapAppToCategory("Outlook"));
        }

        [Fact]
        public void MapAppToCategory_ReturnsDefault_WhenNoMatch()
        {
            var map = new Dictionary<string, TickerCategory>(System.StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = TickerCategory.Warning
            };
            var svc = new WindowsNotificationListenerService(map, new HashSet<string>());

            Assert.Equal(TickerCategory.Warning, svc.MapAppToCategory("UnknownApp"));
        }

        [Fact]
        public void MapAppToCategory_UsesContainsMatch_ForPartialName()
        {
            var map = new Dictionary<string, TickerCategory>(System.StringComparer.OrdinalIgnoreCase)
            {
                ["Teams"] = TickerCategory.Critical,
                ["default"] = TickerCategory.Info
            };
            var svc = new WindowsNotificationListenerService(map, new HashSet<string>());

            // App name "Microsoft Teams" contains "Teams"
            Assert.Equal(TickerCategory.Critical, svc.MapAppToCategory("Microsoft Teams"));
        }

        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        [InlineData(null)]
        public void MapAppToCategory_ReturnsInfo_ForNullOrWhitespaceName(string appName)
        {
            var map = new Dictionary<string, TickerCategory>(System.StringComparer.OrdinalIgnoreCase)
            {
                ["default"] = TickerCategory.Warning
            };
            var svc = new WindowsNotificationListenerService(map, new HashSet<string>());

            Assert.Equal(TickerCategory.Info, svc.MapAppToCategory(appName));
        }
    }
}
