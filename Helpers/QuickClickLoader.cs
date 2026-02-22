using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using NaturalCommands.Models;

namespace NaturalCommands.Helpers
{
    public static class QuickClickLoader
    {
        private static readonly string ConfigPath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "quick_clicks.json"));

        /// <summary>
        /// Loaded profiles (in-memory). Public for read-only access by other components.
        /// </summary>
        public static List<QuickClickProfile> Profiles { get; } = new List<QuickClickProfile>();

        private static JsonSerializerOptions JsonOptions => new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true,
            WriteIndented = true,
            Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) }
        };

        /// <summary>
        /// Load profiles from `quick_clicks.json` if present. Swaps the in-memory list on success.
        /// </summary>
        public static void Load()
        {
            Profiles.Clear();
            try { Logger.LogInfo("QuickClickLoader is disabled."); } catch { }
        }

        /// <summary>
        /// Save provided profiles to disk and update the in-memory copy.
        /// </summary>
        public static void Save(IEnumerable<QuickClickProfile> profiles)
        {
            Profiles.Clear();
            try { Logger.LogInfo("QuickClickLoader.Save ignored because loader is disabled."); } catch { }
        }

        /// <summary>
        /// Returns profiles that match the provided process name and optional window title.
        /// If monitorWidth/monitorHeight are supplied the returned set will include profiles
        /// whose recorded monitor resolution matches the supplied values or profiles that
        /// explicitly target "any monitor" (MonitorWidth/MonitorHeight == null).
        /// </summary>
        public static IEnumerable<QuickClickProfile> GetProfilesForApp(string processName, string? windowTitle, int? monitorWidth = null, int? monitorHeight = null)
        {
            return Enumerable.Empty<QuickClickProfile>();
        }
    }
}
