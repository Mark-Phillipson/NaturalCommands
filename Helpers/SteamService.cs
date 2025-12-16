using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Win32;

namespace NaturalCommands.Helpers
{
    // Minimal Steam detection helpers for installed games (Steam-only)
    public static class SteamService
    {
        public record SteamGame(string AppId, string Name);

        // Attempts to locate the Steam install path from registry or common default
        public static string? GetSteamPath()
        {
            try
            {
                var reg = Registry.GetValue("HKEY_CURRENT_USER\\Software\\Valve\\Steam", "SteamPath", null);
                if (reg is string s && Directory.Exists(s)) return s;
            }
            catch { }

            // Common default
            var defaultPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Steam");
            if (Directory.Exists(defaultPath)) return defaultPath;
            return null;
        }

        // Returns all installed Steam games by scanning steamapps/appmanifest_*.acf files
        public static IEnumerable<SteamGame> GetInstalledGames(string? steamPath = null)
        {
            var results = new List<SteamGame>();
            steamPath ??= GetSteamPath();
            if (string.IsNullOrEmpty(steamPath) || !Directory.Exists(steamPath)) return results;

            var libraryPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            // default library
            libraryPaths.Add(Path.Combine(steamPath, "steamapps"));

            try
            {
                var libFile = Path.Combine(steamPath, "steamapps", "libraryfolders.vdf");
                if (File.Exists(libFile))
                {
                    var text = File.ReadAllText(libFile);
                    static string UnescapeVdfPath(string s)
                    {
                        if (string.IsNullOrWhiteSpace(s)) return string.Empty;
                        // VDF commonly escapes backslashes like D:\\SteamLibrary
                        return s.Replace("\\\\", "\\").Trim();
                    }

                    // Modern format (Steam client): nested blocks with explicit "path" keys
                    // Example:
                    // "1" { "path" "D:\\SteamLibrary" ... }
                    foreach (Match m in Regex.Matches(text, "\\\"path\\\"\\s*\\\"([^\\\"]+)\\\"", RegexOptions.IgnoreCase))
                    {
                        var raw = m.Groups[1].Value;
                        var parent = UnescapeVdfPath(raw);
                        if (string.IsNullOrWhiteSpace(parent)) continue;

                        // Some VDFs might already include steamapps; normalize to steamapps folder
                        var candidate = parent.EndsWith("steamapps", StringComparison.OrdinalIgnoreCase)
                            ? parent
                            : Path.Combine(parent, "steamapps");

                        if (Directory.Exists(candidate)) libraryPaths.Add(candidate);
                    }

                    // Older format: numeric keys map directly to library root paths
                    // Example:
                    // "1" "D:\\SteamLibrary"
                    foreach (Match m in Regex.Matches(text, "\\\"\\d+\\\"\\s*\\\"([^\\\"]+)\\\"", RegexOptions.IgnoreCase))
                    {
                        var raw = m.Groups[1].Value;
                        var parent = UnescapeVdfPath(raw);
                        if (string.IsNullOrWhiteSpace(parent)) continue;
                        if (parent.Equals("path", StringComparison.OrdinalIgnoreCase)) continue;

                        var candidate = parent.EndsWith("steamapps", StringComparison.OrdinalIgnoreCase)
                            ? parent
                            : Path.Combine(parent, "steamapps");

                        if (Directory.Exists(candidate)) libraryPaths.Add(candidate);
                    }

                    // Back-compat: if any quoted string already contains a steamapps path, keep it.
                    foreach (Match m in Regex.Matches(text, "\\\"([^\\\"]*steamapps[^\\\"]*)\\\"", RegexOptions.IgnoreCase))
                    {
                        var p = UnescapeVdfPath(m.Groups[1].Value);
                        if (!string.IsNullOrWhiteSpace(p) && Directory.Exists(p))
                            libraryPaths.Add(p);
                    }
                }
            }
            catch { }

            foreach (var lib in libraryPaths)
            {
                try
                {
                    var dir = new DirectoryInfo(lib);
                    if (!dir.Exists) continue;
                    foreach (var file in dir.GetFiles("appmanifest_*.acf"))
                    {
                        var content = File.ReadAllText(file.FullName);
                        var appid = Regex.Match(content, "\"appid\"\\s*\"(\\d+)\"").Groups[1].Value;
                        var name = Regex.Match(content, "\"name\"\\s*\"([^\"]+)\"").Groups[1].Value;
                        if (!string.IsNullOrEmpty(appid) && !string.IsNullOrEmpty(name))
                        {
                            results.Add(new SteamGame(appid, name));
                        }
                    }
                }
                catch { }
            }

            return results;
        }

        // Simple name-based lookup (more robust normalization)
        public static SteamGame? FindGameByName(string name, string? steamPath = null)
        {
            if (string.IsNullOrWhiteSpace(name)) return null;
            var candidates = GetInstalledGames(steamPath).ToList();

            string NormalizeForLookup(string s)
            {
                if (string.IsNullOrWhiteSpace(s)) return string.Empty;
                s = s.ToLowerInvariant();
                // normalize common spelled numbers/words
                s = Regex.Replace(s, "\\b(two|to)\\b", "2");
                // remove non-alphanumeric characters
                s = Regex.Replace(s, "[^a-z0-9]", "");
                return s;
            }

            var normInput = NormalizeForLookup(name);
            if (string.IsNullOrEmpty(normInput)) return null;

            // Exact name match (case-insensitive)
            var exact = candidates.FirstOrDefault(g => string.Equals(g.Name, name, StringComparison.OrdinalIgnoreCase));
            if (exact != null) return exact;

            // Normalized equality
            var normExact = candidates.FirstOrDefault(g => NormalizeForLookup(g.Name) == normInput);
            if (normExact != null) return normExact;

            // Normalized substring match
            var sub = candidates.FirstOrDefault(g => NormalizeForLookup(g.Name).Contains(normInput));
            if (sub != null) return sub;

            // Fallback: normalized startswith
            return candidates.FirstOrDefault(g => NormalizeForLookup(g.Name).StartsWith(normInput));
        }
    }
}
