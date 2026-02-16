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
            try
            {
                Profiles.Clear();
                try { Logger.LogDebug($"QuickClickLoader.ConfigPath: {ConfigPath}"); } catch { }

                if (!File.Exists(ConfigPath))
                {
                    try { Logger.LogWarning($"QuickClickLoader: config not found at {ConfigPath}"); } catch { }
                    return;
                }

                var json = File.ReadAllText(ConfigPath);
                var loaded = JsonSerializer.Deserialize<List<QuickClickProfile>>(json, JsonOptions);
                if (loaded != null)
                {
                    Profiles.AddRange(loaded);
                    try { Logger.LogDebug($"QuickClickLoader: loaded {Profiles.Count} profile(s)"); } catch { }
                }
            }
            catch (Exception ex)
            {
                try { Logger.LogError($"QuickClickLoader.Load failed: {ex.Message}"); } catch { }
            }
        }

        /// <summary>
        /// Save provided profiles to disk and update the in-memory copy.
        /// </summary>
        public static void Save(IEnumerable<QuickClickProfile> profiles)
        {
            try
            {
                var asList = profiles?.ToList() ?? new List<QuickClickProfile>();
                var json = JsonSerializer.Serialize(asList, JsonOptions);
                File.WriteAllText(ConfigPath, json);

                Profiles.Clear();
                Profiles.AddRange(asList);
                try { Logger.LogDebug($"QuickClickLoader: saved {Profiles.Count} profile(s) to {ConfigPath}"); } catch { }

                // regenerate Talon files after saving
                try { Helpers.QuickClickTalonGenerator.GenerateFromProfiles(Profiles); } catch (Exception ex) { try { Logger.LogError($"QuickClickLoader: talon generation failed: {ex.Message}"); } catch { } }
            }
            catch (Exception ex)
            {
                try { Logger.LogError($"QuickClickLoader.Save failed: {ex.Message}"); } catch { }
                throw;
            }
        }

        /// <summary>
        /// Returns profiles that match the provided process name and optional window title.
        /// If monitorWidth/monitorHeight are supplied the returned set will include profiles
        /// whose recorded monitor resolution matches the supplied values or profiles that
        /// explicitly target "any monitor" (MonitorWidth/MonitorHeight == null).
        /// </summary>
        public static IEnumerable<QuickClickProfile> GetProfilesForApp(string processName, string? windowTitle, int? monitorWidth = null, int? monitorHeight = null)
        {
            var results = new List<QuickClickProfile>();
            if (string.IsNullOrWhiteSpace(processName)) return results;

            foreach (var p in Profiles)
            {
                try
                {
                    if (!string.Equals(p.ProcessName, processName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    // Window title matching (null/empty = match any)
                    if (!string.IsNullOrWhiteSpace(p.WindowTitlePattern))
                    {
                        if (string.IsNullOrWhiteSpace(windowTitle))
                            continue;

                        var pattern = p.WindowTitlePattern!;
                        bool titleMatches = false;

                        // Try regex first (safe guarded)
                        try
                        {
                            if (Regex.IsMatch(windowTitle, pattern, RegexOptions.IgnoreCase))
                                titleMatches = true;
                        }
                        catch
                        {
                            // fallback to substring match when regex fails
                            if (windowTitle.IndexOf(pattern, StringComparison.OrdinalIgnoreCase) >= 0)
                                titleMatches = true;
                        }

                        if (!titleMatches) continue;
                    }

                    // Monitor resolution matching: if monitor args provided, require either an exact match
                    // or a profile with null MonitorWidth/MonitorHeight (applies to any monitor).
                    if (monitorWidth.HasValue && monitorHeight.HasValue)
                    {
                        if (p.MonitorWidth.HasValue && p.MonitorHeight.HasValue)
                        {
                            if (p.MonitorWidth.Value != monitorWidth.Value || p.MonitorHeight.Value != monitorHeight.Value)
                                continue; // resolution-specific profile that doesn't match
                        }
                        // else: profile applies to any monitor -> accept
                    }

                    results.Add(p);
                }
                catch { /* individual-profile errors shouldn't break the enumeration */ }
            }

            return results;
        }
    }
}
