using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace NaturalCommands
{
    /// <summary>
    /// Centralized manager that ensures exactly one persistent TickerOverlayForm
    /// exists per process and provides a thread-safe API to add messages.
    /// </summary>
    internal static class TickerOverlayManager
    {
        private static TickerOverlayForm? _form;
        private static Thread? _uiThread;
        private static ManualResetEventSlim? _handleReady;
        private static readonly object _lock = new();

        /// <summary>
        /// Ensure a persistent ticker form exists and seed it with the provided lines.
        /// If the form already exists, the lines are appended instead of recreating.
        /// </summary>
        public static void EnsurePersistentTicker(IEnumerable<string> initialLines, int cycleSeconds = 5, int maxCycles = 0, bool topPosition = false, bool hideOnDismiss = true)
        {
            if (initialLines == null) initialLines = Array.Empty<string>();

            lock (_lock)
            {
                // If an active form exists, append the lines and return
                if (_form != null && !_form.IsDisposed)
                {
                    try
                    {
                        foreach (var m in initialLines.Select(TickerOverlayForm.ParseSingleLine))
                        {
                            _form.AddMessage(m);
                        }
                    }
                    catch { }

                    return;
                }

                // If UI thread already started, wait briefly for handle
                if (_uiThread != null && _uiThread.IsAlive && _handleReady != null && !_handleReady.IsSet)
                {
                    try { _handleReady.Wait(TimeSpan.FromSeconds(5)); } catch { }
                    if (_form != null && !_form.IsDisposed)
                    {
                        try
                        {
                            foreach (var m in initialLines.Select(TickerOverlayForm.ParseSingleLine))
                                _form.AddMessage(m);
                        }
                        catch { }
                        return;
                    }
                }

                // Otherwise create a new STA UI thread that owns the form
                _handleReady?.Dispose();
                _handleReady = new ManualResetEventSlim(false);

                var linesCopy = initialLines.ToList();

                _uiThread = new Thread(() =>
                {
                    try
                    {
                        Application.EnableVisualStyles();
                        Application.SetCompatibleTextRenderingDefault(false);

                        var form = new TickerOverlayForm(linesCopy, cycleSeconds: cycleSeconds, maxCycles: maxCycles, topPosition: topPosition, hideOnDismiss: hideOnDismiss);
                        form.FormClosed += (s, e) => { lock (_lock) { _form = null; } };

                        // Force handle creation so callers can Invoke safely
                        var _ = form.Handle;

                        lock (_lock) { _form = form; }

                        try { _handleReady.Set(); } catch { }

                        try { Helpers.Logger.LogInfo("[TICKER-MGR] Persistent ticker created"); } catch { }

                        Application.Run(form);
                    }
                    catch (Exception ex)
                    {
                        try { Helpers.Logger.LogError($"[TICKER-MGR] Failed to create ticker: {ex.Message}"); } catch { }
                    }
                    finally
                    {
                        lock (_lock)
                        {
                            try { _handleReady?.Dispose(); } catch { }
                            _handleReady = null;
                            _uiThread = null;
                            _form = null;
                        }
                    }
                })
                {
                    IsBackground = false
                };

                _uiThread.SetApartmentState(ApartmentState.STA);
                _uiThread.Start();

                try { _handleReady.Wait(TimeSpan.FromSeconds(5)); } catch { }
            }
        }

        /// <summary>
        /// Add messages to the persistent ticker. Ensures a persistent form exists.
        /// </summary>
        public static void AddMessages(IEnumerable<TickerMessage> messages)
        {
            if (messages == null) return;

            lock (_lock)
            {
                if (_form == null || _form.IsDisposed)
                {
                    // Create a persistent ticker with a placeholder so we have a UI owner
                    EnsurePersistentTicker(new[] { "Ticker ready" }, cycleSeconds: 5, maxCycles: 0, topPosition: false, hideOnDismiss: true);
                }

                if (_form != null && !_form.IsDisposed)
                {
                    foreach (var msg in messages)
                    {
                        try { _form.AddMessage(msg); } catch (Exception ex) { try { Helpers.Logger.LogWarning($"[TICKER-MGR] Failed to enqueue message: {ex.Message}"); } catch { } }
                    }
                }
            }
        }

        public static bool IsAlive
        {
            get { lock (_lock) { return _form != null && !_form.IsDisposed; } }
        }
    }
}
