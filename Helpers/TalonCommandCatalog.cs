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
        private static bool s_loaded;

        public static IReadOnlyList<string> GetCommands()
        {
            EnsureLoaded();
            return s_commands;
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
            if (Directory.Exists(scriptRoot))
            {
                try
                {
                    foreach (var file in Directory.EnumerateFiles(scriptRoot, "*.talon", SearchOption.AllDirectories))
                    {
                        ParseTalonFile(file, commands);
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

            if (s_commands.Count > 0)
            {
                TryWriteCache(cachePath, s_commands);
            }

            Logger.LogInfo($"TalonCommandCatalog: Loaded {s_commands.Count} Talon commands.");
        }

        private static void ParseTalonFile(string filePath, HashSet<string> commands)
        {
            try
            {
                var lines = File.ReadAllLines(filePath);
                bool inCommandBlock = false;
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

                    if (!trimmed.EndsWith(":", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    var phrase = trimmed.Substring(0, trimmed.Length - 1).Trim();
                    if (string.IsNullOrWhiteSpace(phrase))
                    {
                        continue;
                    }

                    if (phrase.Contains("=", StringComparison.Ordinal) || phrase.StartsWith("tag(", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    commands.Add(phrase);
                }
            }
            catch (Exception ex)
            {
                Logger.LogWarning($"TalonCommandCatalog: Failed parsing '{filePath}': {ex.Message}");
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
    }
}
