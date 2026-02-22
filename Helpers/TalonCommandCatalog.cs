using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace NaturalCommands.Helpers
{
    public static class TalonCommandCatalog
    {
        private static readonly object s_lock = new object();
        private static List<string> s_commands = new List<string>();
        private static Dictionary<string, double> s_contextPreference = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
        private static bool s_loaded;

        public static IReadOnlyList<string> GetCommands()
        {
            EnsureLoaded();
            return s_commands;
        }

        public static IReadOnlyDictionary<string, double> GetContextPreferences()
        {
            EnsureLoaded();
            return s_contextPreference;
        }

        public static void EnsureLoaded()
        {
            if (s_loaded) return;
            lock (s_lock)
            {
                if (s_loaded) return;
                LoadInternal();
                s_loaded = true;
            }
        }

        public static int Reload()
        {
            lock (s_lock)
            {
                LoadInternal();
                s_loaded = true;
                return s_commands.Count;
            }
        }

        private static void LoadInternal()
        {
            var settings = Models.AppSettings.Instance.Talon;
            var scriptRoot = settings.UserScriptsDirectory ?? string.Empty;
            var cachePath = GetCachePath(settings.CatalogCacheFileName);

            var commands = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var commandStats = new Dictionary<string, CommandStats>(StringComparer.OrdinalIgnoreCase);
            if (Directory.Exists(scriptRoot))
            {
                try
                {
                    foreach (var file in Directory.EnumerateFiles(scriptRoot, "*.talon", SearchOption.AllDirectories))
                    {
                        ParseTalonFile(file, commands, commandStats);
                    }
                }
                catch (Exception ex)
                {
                    Logger.LogWarning($"TalonCommandCatalog: Failed scanning scripts under '{scriptRoot}': {ex.Message}");
                }
            }

            if (commands.Count == 0)
            {
                TryLoadFromCache(cachePath, commands);
            }

            s_commands = commands
                .Where(c => !string.IsNullOrWhiteSpace(c))
                .Select(c => c.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(c => c, StringComparer.OrdinalIgnoreCase)
                .ToList();

            s_contextPreference = BuildContextPreference(s_commands, commandStats);

            if (s_commands.Count > 0)
            {
                TryWriteCache(cachePath, s_commands);
            }

            Logger.LogInfo($"TalonCommandCatalog: Loaded {s_commands.Count} Talon commands.");
        }

        private static void ParseTalonFile(string filePath, HashSet<string> commands, Dictionary<string, CommandStats> commandStats)
        {
            try
            {
                var lines = File.ReadAllLines(filePath);
                bool hasCommandDelimiter = lines.Any(l => string.Equals(l.Trim(), "-", StringComparison.Ordinal));
                bool hasAppContext = false;
                if (hasCommandDelimiter)
                {
                    foreach (var line in lines)
                    {
                        var trimmed = line.Trim();
                        if (trimmed == "-")
                        {
                            break;
                        }

                        var lowerContext = trimmed.ToLowerInvariant();
                        if (lowerContext.StartsWith("app:")
                            || lowerContext.StartsWith("app.exe:")
                            || lowerContext.StartsWith("title:"))
                        {
                            hasAppContext = true;
                            break;
                        }
                    }
                }

                bool inCommandBlock = !hasCommandDelimiter;
                foreach (var rawLine in lines)
                {
                    var trimmed = rawLine.Trim();
                    if (string.IsNullOrWhiteSpace(trimmed) || trimmed.StartsWith("#"))
                    {
                        continue;
                    }

                    if (trimmed == "-")
                    {
                        inCommandBlock = true;
                        continue;
                    }

                    if (!inCommandBlock)
                    {
                        continue;
                    }

                    if (rawLine.StartsWith(" ") || rawLine.StartsWith("\t"))
                    {
                        continue;
                    }

                    var colonIndex = trimmed.IndexOf(':');
                    if (colonIndex <= 0)
                    {
                        continue;
                    }

                    var lower = trimmed.ToLowerInvariant();
                    if (lower.StartsWith("os:")
                        || lower.StartsWith("app:")
                        || lower.StartsWith("app.exe:")
                        || lower.StartsWith("mode:")
                        || lower.StartsWith("tag:")
                        || lower.StartsWith("title:")
                        || lower.StartsWith("hostname:")
                        || lower.StartsWith("and "))
                    {
                        continue;
                    }

                    var phrase = trimmed.Substring(0, colonIndex).Trim();
                    if (string.IsNullOrWhiteSpace(phrase))
                    {
                        continue;
                    }

                    if (phrase.Contains('^') || phrase.Contains('$'))
                    {
                        continue;
                    }

                    if (phrase.Contains("=", StringComparison.Ordinal) || phrase.StartsWith("tag(", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    commands.Add(phrase);
                    AddCommandStats(commandStats, phrase, hasAppContext);
                }
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"TalonCommandCatalog: Failed parsing '{filePath}': {ex.Message}");
            }
        }

        private static Dictionary<string, double> BuildContextPreference(List<string> commands, Dictionary<string, CommandStats> commandStats)
        {
            var preference = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            foreach (var command in commands)
            {
                if (!commandStats.TryGetValue(command, out var stats))
                {
                    preference[command] = 0.5;
                    continue;
                }

                var total = stats.GlobalCount + stats.AppScopedCount;
                if (total <= 0)
                {
                    preference[command] = 0.5;
                    continue;
                }

                preference[command] = (stats.GlobalCount + 0.25) / (total + 0.5);
            }

            return preference;
        }

        private static void AddCommandStats(Dictionary<string, CommandStats> commandStats, string command, bool isAppScoped)
        {
            if (!commandStats.TryGetValue(command, out var stats))
            {
                stats = new CommandStats();
                commandStats[command] = stats;
            }

            if (isAppScoped)
            {
                stats.AppScopedCount++;
            }
            else
            {
                stats.GlobalCount++;
            }
        }

        private static string GetCachePath(string? cacheFileName)
        {
            var fileName = string.IsNullOrWhiteSpace(cacheFileName) ? "talon_phrase_catalog.json" : cacheFileName.Trim();
            return Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, fileName));
        }

        private static void TryLoadFromCache(string cachePath, HashSet<string> commands)
        {
            try
            {
                if (!File.Exists(cachePath)) return;
                var json = File.ReadAllText(cachePath);
                var payload = JsonSerializer.Deserialize<TalonCommandCatalogCache>(json);
                if (payload?.Commands == null) return;
                foreach (var command in payload.Commands)
                {
                    if (!string.IsNullOrWhiteSpace(command))
                    {
                        commands.Add(command.Trim());
                    }
                }
                Logger.LogInfo($"TalonCommandCatalog: Loaded {commands.Count} commands from cache.");
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"TalonCommandCatalog: Failed loading cache '{cachePath}': {ex.Message}");
            }
        }

        private static void TryWriteCache(string cachePath, List<string> commands)
        {
            try
            {
                var dir = Path.GetDirectoryName(cachePath);
                if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                var payload = new TalonCommandCatalogCache
                {
                    GeneratedUtc = DateTime.UtcNow,
                    Commands = commands
                };

                var json = JsonSerializer.Serialize(payload, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(cachePath, json);
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"TalonCommandCatalog: Failed writing cache '{cachePath}': {ex.Message}");
            }
        }

        private sealed class TalonCommandCatalogCache
        {
            public DateTime GeneratedUtc { get; set; }
            public List<string> Commands { get; set; } = new List<string>();
        }

        private sealed class CommandStats
        {
            public int GlobalCount { get; set; }
            public int AppScopedCount { get; set; }
        }
    }
}
