using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace NaturalCommands.Helpers
{
    public static class MultiActionLoader
    {
        public static Dictionary<string, NaturalCommands.RunMultipleActionsAction> Commands { get; } = new(StringComparer.OrdinalIgnoreCase);

        private static readonly string ConfigPath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "multi_actions.json"));

        public static void Load()
        {
            try
            {
                Commands.Clear();
                var logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "app.log");
                try { NaturalCommands.Helpers.Logger.LogDebug($"MultiActionLoader.ConfigPath: {ConfigPath}"); } catch { }
                if (!File.Exists(ConfigPath))
                {
                    try { NaturalCommands.Helpers.Logger.LogWarning($"MultiActionLoader: config not found at {ConfigPath}"); } catch { }
                    return;
                }
                var json = File.ReadAllText(ConfigPath);
                var doc = JsonDocument.Parse(json);
                if (doc.RootElement.ValueKind != JsonValueKind.Array) return;
                var registered = new List<string>();
                foreach (var item in doc.RootElement.EnumerateArray())
                {
                    try
                    {
                        var name = item.GetProperty("Name").GetString() ?? string.Empty;
                        bool continueOnError = true;
                        int delay = 250;
                        if (item.TryGetProperty("ContinueOnError", out var pco)) continueOnError = pco.GetBoolean();
                        if (item.TryGetProperty("DelayMsBetween", out var pd)) delay = pd.GetInt32();

                        var actions = new List<NaturalCommands.ActionBase>();
                        if (item.TryGetProperty("Actions", out var aProp) && aProp.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var a in aProp.EnumerateArray())
                            {
                                var act = DeserializeAction(a);
                                if (act != null) actions.Add(act);
                            }
                        }

                        if (!string.IsNullOrWhiteSpace(name) && actions.Count > 0)
                        {
                            var r = new NaturalCommands.RunMultipleActionsAction(name, actions, continueOnError, delay);
                            // store the entry under its original name
                            Commands[name] = r;
                            registered.Add(name);
                            // also store a normalized/compact key to allow flexible matching (e.g. "setup" vs "set up")
                            try
                            {
                                var normalized = NormalizeKey(name);
                                if (!string.IsNullOrWhiteSpace(normalized) && !Commands.ContainsKey(normalized))
                                {
                                    Commands[normalized] = r;
                                    registered.Add(normalized);
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }
                }
            }
            catch { }
            try
            {
                var logPath2 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "app.log");
                // write keys that were registered (if any)
                if (Commands.Count > 0)
                {
                    try { NaturalCommands.Helpers.Logger.LogDebug($"MultiActionLoader loaded keys: {string.Join(", ", Commands.Keys)}"); } catch { }
                }
            }
            catch { }
        }

        public static string NormalizeKey(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
            try
            {
                var low = s.ToLowerInvariant();
                // remove punctuation (keep word chars and whitespace)
                low = System.Text.RegularExpressions.Regex.Replace(low, "[^\\w\\s]", "");
                // collapse whitespace
                low = System.Text.RegularExpressions.Regex.Replace(low, "\\s+", " ").Trim();
                // compact (remove spaces) so variants like "setup" and "set up" match
                var compact = low.Replace(" ", "");
                return compact;
            }
            catch
            {
                return s.ToLowerInvariant();
            }
        }

        private static NaturalCommands.ActionBase? DeserializeAction(JsonElement elem)
        {
            if (elem.ValueKind != JsonValueKind.Object) return null;
            string type = elem.GetProperty("Type").GetString() ?? string.Empty;
            try
            {
                switch (type)
                {
                    case "SendKeysAction":
                        return new NaturalCommands.SendKeysAction(elem.GetProperty("KeysText").GetString() ?? string.Empty);
                    case "LaunchAppAction":
                        return new NaturalCommands.LaunchAppAction(elem.GetProperty("AppExe").GetString() ?? string.Empty);
                    case "OpenWebsiteAction":
                        return new NaturalCommands.OpenWebsiteAction(elem.GetProperty("Url").GetString() ?? string.Empty);
                    case "OpenFolderAction":
                        return new NaturalCommands.OpenFolderAction(elem.GetProperty("KnownFolder").GetString() ?? string.Empty);
                    case "FocusWindowAction":
                        return new NaturalCommands.FocusWindowAction(elem.GetProperty("WindowTitleSubstring").GetString() ?? string.Empty);
                    case "ExecuteVSCommandAction":
                        return new NaturalCommands.ExecuteVSCommandAction(elem.GetProperty("CommandName").GetString() ?? string.Empty);
                    case "MoveWindowAction":
                        var target = elem.GetProperty("Target").GetString() ?? "active";
                        var monitor = elem.GetProperty("Monitor").GetString() ?? "current";
                        string? pos = null;
                        int? w = null; int? h = null;
                        if (elem.TryGetProperty("Position", out var pp) && pp.ValueKind == JsonValueKind.String) pos = pp.GetString();
                        if (elem.TryGetProperty("WidthPercent", out var wp) && wp.ValueKind == JsonValueKind.Number) w = wp.GetInt32();
                        if (elem.TryGetProperty("HeightPercent", out var hp) && hp.ValueKind == JsonValueKind.Number) h = hp.GetInt32();
                        return new NaturalCommands.MoveWindowAction(target, monitor, pos, w, h);
                    default:
                        return null;
                }
            }
            catch
            {
                return null;
            }
        }
    }
}
