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
                    NaturalCommands.Helpers.Logger.LogError($"[NotificationListener] RequestAccess failed: {ex}");
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
                                // AppInfo cast throws for unpackaged/Win32 apps — guard separately
                                string appName;
                                try { appName = notification.AppInfo?.DisplayInfo?.DisplayName ?? "Unknown"; }
                                catch { appName = "Unknown"; }

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

                        // Show all notifications from this cycle in a single form
                        if (newLines.Count > 0)
                            ShowTicker(newLines);
                    }
                    catch (ObjectDisposedException ex)
                    {
                        NaturalCommands.Helpers.Logger.LogError($"[NotificationListener] Listener disposed: {ex}");
                        Console.WriteLine("[listen-notifications] Listener disposed — attempting reinitialization.");

                        bool reinitSuccess = false;
                        for (int attempt = 0; attempt < 5 && !token.IsCancellationRequested; attempt++)
                        {
                            try
                            {
                                Thread.Sleep(1000);
                                listener = UserNotificationListener.Current;
                                // Probe the listener to ensure it's usable
                                var probe = listener.GetNotificationsAsync(NotificationKinds.Toast).GetAwaiter().GetResult();
                                reinitSuccess = true;
                                NaturalCommands.Helpers.Logger.LogInfo($"[NotificationListener] Reinitialized listener after {attempt + 1} attempt(s).");
                                break;
                            }
                            catch (Exception rex)
                            {
                                NaturalCommands.Helpers.Logger.LogWarning($"[NotificationListener] Reinit attempt {attempt + 1} failed: {rex.Message}");
                            }
                        }

                        if (!reinitSuccess)
                        {
                            NaturalCommands.Helpers.Logger.LogError("[NotificationListener] Reinit failed — stopping poll loop.");
                            Console.WriteLine("[listen-notifications] Listener reinit failed, exiting.");
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

        private TickerOverlayForm? _activeForm;
        private readonly object _formLock = new();

        private void ShowTicker(System.Collections.Generic.List<string> lines)
        {
            lock (_formLock)
            {
                // Active form exists — add all new messages to it via Invoke
                if (_activeForm != null && !_activeForm.IsDisposed)
                {
                    bool invokeOk = true;
                    foreach (var line in lines)
                    {
                        try { _activeForm.AddMessage(TickerOverlayForm.ParseSingleLine(line)); }
                        catch { invokeOk = false; break; }
                    }
                    if (invokeOk) return;
                    // Invoke failed (form closing) — fall through to create new form below
                    _activeForm = null;
                }

                // No active form — create one entirely on a new STA thread so the handle
                // is owned by that thread (required for Invoke to work on future AddMessage calls)
                var linesCopy = new System.Collections.Generic.List<string>(lines);
                TickerOverlayForm? created = null;
                var handleReady = new ManualResetEventSlim(false);

                var thread = new Thread(() =>
                {
                    System.Windows.Forms.Application.EnableVisualStyles();
                    System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(false);
                    // maxCycles=0 → never auto-close; user must press Dismiss
                    var form = new TickerOverlayForm(linesCopy, cycleSeconds: 5, maxCycles: 0, topPosition: false);
                    form.FormClosed += (s, e) => { lock (_formLock) { _activeForm = null; } };
                    // Force handle creation on this thread before signalling
                    var _ = form.Handle;
                    created = form;
                    handleReady.Set();
                    System.Windows.Forms.Application.Run(form);
                });
                thread.SetApartmentState(ApartmentState.STA);
                thread.IsBackground = true;
                thread.Start();

                // Wait for handle to be ready so Invoke works for subsequent AddMessage calls
                handleReady.Wait(TimeSpan.FromSeconds(3));
                _activeForm = created;
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
