using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using Windows.UI.Notifications;
using Windows.UI.Notifications.Management;

namespace NaturalCommands
{
    internal sealed class WindowsNotificationListenerService
    {
        private readonly Dictionary<string, TickerCategory> _appCategoryMap;
        private readonly HashSet<string> _ignoreApps;

        public WindowsNotificationListenerService()
        {
            _appCategoryMap = LoadCategoryRules();
            _ignoreApps = LoadIgnoreRules();
        }

        /// <summary>
        /// Test-only constructor: injects pre-built rule maps so that WinRT is never touched.
        /// </summary>
        internal WindowsNotificationListenerService(
            Dictionary<string, TickerCategory> appCategoryMap,
            HashSet<string> ignoreApps)
        {
            _appCategoryMap = appCategoryMap;
            _ignoreApps = ignoreApps;
        }

        private static Dictionary<string, TickerCategory> LoadCategoryRules()
        {
            var filePath = Path.Combine(AppContext.BaseDirectory, "notification-category-rules.json");
            if (!File.Exists(filePath))
            {
                return new Dictionary<string, TickerCategory>(StringComparer.OrdinalIgnoreCase)
                {
                    ["default"] = TickerCategory.Info
                };
            }

            try
            {
                var json = File.ReadAllText(filePath);
                return ParseCategoryRulesJson(json);
            }
            catch
            {
                return new Dictionary<string, TickerCategory>(StringComparer.OrdinalIgnoreCase)
                {
                    ["default"] = TickerCategory.Info
                };
            }
        }

        /// <summary>
        /// Parses a JSON string (array of {app, category} or {default} objects) into a category map.
        /// Exposed as <c>internal</c> so unit tests can exercise the parsing logic directly.
        /// </summary>
        internal static Dictionary<string, TickerCategory> ParseCategoryRulesJson(string json)
        {
            var raw = JsonSerializer.Deserialize<List<Dictionary<string, string>>>(json);
            if (raw == null)
            {
                return new Dictionary<string, TickerCategory>(StringComparer.OrdinalIgnoreCase)
                {
                    ["default"] = TickerCategory.Info
                };
            }

            var map = new Dictionary<string, TickerCategory>(StringComparer.OrdinalIgnoreCase);
            foreach (var item in raw)
            {
                if (item.TryGetValue("app", out var app) && item.TryGetValue("category", out var categoryText))
                {
                    if (Enum.TryParse<TickerCategory>(categoryText, true, out var category))
                    {
                        map[app] = category;
                    }
                }
                else if (item.TryGetValue("default", out var defaultCategory) && Enum.TryParse<TickerCategory>(defaultCategory, true, out var defaultCat))
                {
                    map["default"] = defaultCat;
                }
            }

            if (!map.ContainsKey("default"))
            {
                map["default"] = TickerCategory.Info;
            }

            return map;
        }

        private static HashSet<string> LoadIgnoreRules()
        {
            var filePath = Path.Combine(AppContext.BaseDirectory, "notification-ignore-rules.json");
            if (!File.Exists(filePath))
            {
                return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }

            try
            {
                var json = File.ReadAllText(filePath);
                return ParseIgnoreRulesJson(json);
            }
            catch
            {
                return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }
        }

        /// <summary>
        /// Parses a JSON string (array of {app} objects) into a set of ignored app names.
        /// Exposed as <c>internal</c> so unit tests can exercise the parsing logic directly.
        /// </summary>
        internal static HashSet<string> ParseIgnoreRulesJson(string json)
        {
            var raw = JsonSerializer.Deserialize<List<Dictionary<string, string>>>(json);
            if (raw == null) return new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            return raw
                .Where(i => i.TryGetValue("app", out _))
                .Select(i => i["app"])
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
        }

        private readonly HashSet<uint> _seen = new();
        private CancellationTokenSource? _pollCts;
        public event Action<System.Collections.Generic.List<string>>? NotificationsReceived;
        public event Action? ListenerFailed;

        /// <summary>
        /// Starts a dedicated STA thread that owns ALL WinRT access:
        /// RequestAccess + polling. Returns true if access was granted.
        /// Blocks until the access check completes (max 10 s).
        /// </summary>
        public bool Start()
        {
            _pollCts = new CancellationTokenSource();
            var token = _pollCts.Token;

            bool? accessResult = null;
            var accessKnown = new ManualResetEventSlim(false);

            var thread = new Thread(() =>
            {
                // --- access check on this STA thread ---
                try
                {
                    var status = UserNotificationListener.Current
                        .RequestAccessAsync().GetAwaiter().GetResult();
                    accessResult = status == UserNotificationListenerAccessStatus.Allowed;
                }
                catch (Exception ex)
                {
                    if (ex is System.Runtime.InteropServices.COMException cex)
                    {
                        NaturalCommands.Helpers.Logger.LogError($"[NotificationListener] RequestAccess failed: {cex.Message} (HR: 0x{cex.HResult:X8}). This may indicate the notification manager or COM server is unavailable.");
                    }
                    else
                    {
                        NaturalCommands.Helpers.Logger.LogError($"[NotificationListener] RequestAccess failed: {ex.Message}");
                    }
                    accessResult = false;
                }
                finally
                {
                    accessKnown.Set();
                }

                if (accessResult != true) return;

                // --- poll loop on the same STA thread ---
                var listener = UserNotificationListener.Current;
                NaturalCommands.Helpers.Logger.LogInfo("[NotificationListener] Poll loop started.");
                Console.WriteLine("[listen-notifications] Poll loop started.");

                while (!token.IsCancellationRequested)
                {
                    try
                    {
                        var notifications = listener
                            .GetNotificationsAsync(NotificationKinds.Toast)
                            .GetAwaiter().GetResult();

                        // Collect all new notifications from this poll cycle into one batch
                        var newLines = new System.Collections.Generic.List<string>();
                        foreach (var notification in notifications)
                        {
                            if (!_seen.Add(notification.Id)) continue;
                            try
                            {
                                // AppInfo may be null for unpackaged/Win32 apps — guard and log for diagnostics
                                string appName;
                                try
                                {
                                    if (notification.AppInfo != null)
                                    {
                                        appName = notification.AppInfo.DisplayInfo?.DisplayName ?? "Unknown";
                                    }
                                    else
                                    {
                                        appName = "Unknown";
                                        NaturalCommands.Helpers.Logger.LogInfo($"[NotificationListener] AppInfo is null for id={notification.Id}");
                                    }
                                }
                                catch (InvalidCastException icex)
                                {
                                    appName = "Unknown";
                                    NaturalCommands.Helpers.Logger.LogWarning($"[NotificationListener] AppInfo cast failed for id={notification.Id}: {icex.Message}");
                                }
                                catch (Exception ex)
                                {
                                    appName = "Unknown";
                                    NaturalCommands.Helpers.Logger.LogError($"[NotificationListener] Failed to read AppInfo for id={notification.Id}: {ex.Message}");
                                }

                                if (_ignoreApps.Contains(appName)) continue;

                                var binding = notification.Notification.Visual
                                    .GetBinding(KnownNotificationBindings.ToastGeneric);
                                var title = binding?.GetTextElements()?.FirstOrDefault()?.Text ?? appName;
                                var body  = binding?.GetTextElements()?.Skip(1).FirstOrDefault()?.Text ?? string.Empty;
                                var message = string.IsNullOrWhiteSpace(body) ? title : $"{title} — {body}";

                                NaturalCommands.Helpers.Logger.LogInfo($"[NotificationListener] id={notification.Id} app={appName} title={title} body={body} message={message}");
                                Console.WriteLine($"[listen-notifications] {appName}: {message}");

                                var category = MapAppToCategory(appName);
                                newLines.Add($"{category.ToString().ToLowerInvariant()}:{message}");
                                listener.RemoveNotification(notification.Id);
                            }
                            catch (Exception ex)
                            {
                                NaturalCommands.Helpers.Logger.LogError($"[NotificationListener] Process id={notification.Id}: {ex}");
                            }
                        }

                        // Show or forward all notifications from this cycle
                        if (newLines.Count > 0)
                        {
                            try
                            {
                                if (NotificationsReceived != null)
                                {
                                    NaturalCommands.Helpers.Logger.LogInfo($"[NotificationListener] Forwarding {newLines.Count} lines; HasForwarder=True");
                                    NotificationsReceived.Invoke(newLines);
                                }
                                else
                                {
                                    NaturalCommands.Helpers.Logger.LogInfo("[NotificationListener] No forwarder subscribed; showing local ticker.");
                                    ShowTicker(newLines);
                                }
                            }
                            catch (Exception ex)
                            {
                                NaturalCommands.Helpers.Logger.LogError($"[NotificationListener] Error forwarding notifications: {ex}");
                                // Fallback to local display
                                try { ShowTicker(newLines); } catch { }
                            }
                        }
                    }
                    catch (ObjectDisposedException ex)
                    {
                        NaturalCommands.Helpers.Logger.LogError($"[NotificationListener] Listener disposed: {ex.ToString()} | threadId={Thread.CurrentThread.ManagedThreadId} | tokenCancelled={token.IsCancellationRequested} | seen={_seen.Count} | listenerNull={(listener==null)}");
                        Console.WriteLine("[listen-notifications] Listener disposed — attempting reinitialization.");

                        bool reinitSuccess = false;
                        for (int attempt = 0; attempt < 5 && !token.IsCancellationRequested; attempt++)
                        {
                            try
                            {
                                Thread.Sleep(1000);
                                string probeError = null;
                                bool probeOk = TryProbeListenerOnSta(timeoutMs: 5000, out probeError);
                                if (!probeOk)
                                {
                                    NaturalCommands.Helpers.Logger.LogWarning($"[NotificationListener] Reinit attempt {attempt + 1} probe failed: {probeError ?? "unknown"}");
                                    // If the probe indicates the COM server is not connected, abort early
                                    if (!string.IsNullOrEmpty(probeError) && (probeError.Contains("0x800401FD") || probeError.Contains("CO_E_OBJNOTCONNECTED") || probeError.Contains("Object is not connected to server")))
                                    {
                                        NaturalCommands.Helpers.Logger.LogError($"[NotificationListener] Reinit unrecoverable: {probeError}. Aborting reinit.");
                                        break;
                                    }
                                    continue;
                                }

                                // Create a fresh listener proxy on this (polling) STA thread
                                listener = UserNotificationListener.Current;
                                // quick sanity-check probe
                                var check = listener.GetNotificationsAsync(NotificationKinds.Toast).GetAwaiter().GetResult();

                                reinitSuccess = true;
                                NaturalCommands.Helpers.Logger.LogInfo($"[NotificationListener] Reinitialized listener after {attempt + 1} attempt(s).");
                                break;
                            }
                            catch (Exception rex)
                            {
                                NaturalCommands.Helpers.Logger.LogWarning($"[NotificationListener] Reinit attempt {attempt + 1} failed: {rex.ToString()}");
                            }
                        }

                        if (!reinitSuccess)
                        {
                            NaturalCommands.Helpers.Logger.LogError("[NotificationListener] Reinit failed — stopping poll loop.");
                            Console.WriteLine("[listen-notifications] Listener reinit failed, exiting.");
                            try { ListenerFailed?.Invoke(); } catch { }
                            break;
                        }
                        else
                        {
                            continue;
                        }
                    }
                    catch (Exception ex)
                    {
                        NaturalCommands.Helpers.Logger.LogError($"[NotificationListener] Poll error: {ex}");
                    }

                    Thread.Sleep(1000);
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();

            accessKnown.Wait(TimeSpan.FromSeconds(10));
            return accessResult == true;
        }

        public void Stop() => _pollCts?.Cancel();

        private bool TryProbeListenerOnSta(int timeoutMs, out string? probeError)
        {
            probeError = null;
            bool success = false;
            var done = new System.Threading.ManualResetEventSlim(false);
            string localError = null;

            var thread = new Thread(() =>
            {
                try
                {
                    try
                    {
                        var status = UserNotificationListener.Current.RequestAccessAsync().GetAwaiter().GetResult();
                        if (status != UserNotificationListenerAccessStatus.Allowed)
                        {
                            localError = $"RequestAccess returned {status}";
                        }
                        else
                        {
                            var probe = UserNotificationListener.Current.GetNotificationsAsync(NotificationKinds.Toast).GetAwaiter().GetResult();
                            success = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        if (ex is System.Runtime.InteropServices.COMException cex)
                        {
                            localError = $"{cex.Message} (HR: 0x{cex.HResult:X8})";
                        }
                        else
                        {
                            localError = ex.ToString();
                        }
                    }
                }
                finally
                {
                    done.Set();
                }
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.IsBackground = true;
            thread.Start();

            if (!done.Wait(TimeSpan.FromMilliseconds(timeoutMs)))
            {
                probeError = "probe timeout";
                return false;
            }

            probeError = localError;
            return success;
        }

        private void ShowTicker(System.Collections.Generic.List<string> lines)
        {
            try
            {
                // Delegate to centralized manager which ensures a single persistent form
                TickerOverlayManager.EnsurePersistentTicker(lines, cycleSeconds: 5, maxCycles: 0, topPosition: false, hideOnDismiss: false);
            }
            catch (Exception ex)
            {
                try { NaturalCommands.Helpers.Logger.LogError($"[NotificationListener] ShowTicker failed: {ex.Message}"); } catch { }
            }
        }

        internal TickerCategory MapAppToCategory(string? appName)
        {
            if (string.IsNullOrWhiteSpace(appName))
                return TickerCategory.Info;

            if (_appCategoryMap.TryGetValue(appName, out var cat))
                return cat;

            // look for app name contains rule
            var match = _appCategoryMap.FirstOrDefault(kvp => kvp.Key != "default" && appName.IndexOf(kvp.Key, StringComparison.OrdinalIgnoreCase) >= 0);
            if (!match.Equals(default(KeyValuePair<string, TickerCategory>)))
            {
                return match.Value;
            }

            return _appCategoryMap.TryGetValue("default", out var defaultCat) ? defaultCat : TickerCategory.Info;
        }
    }
}
