using System;

namespace NaturalCommands
{
    /// <summary>
    /// Centralized test-only mode helper. Tests and CI can enable test-only mode
    /// by setting the `NATURALCOMMANDS_TEST_MODE` or `NATURALCOMMANDS_TEST_NO_LAUNCH` environment
    /// variable, or by setting `AppSettings.Instance.Behavior.TestMode = true` in a test.
    /// </summary>
    public static class TestMode
    {
        private static bool? s_override;

        public static bool IsEnabled
        {
            get
            {
                if (s_override.HasValue) return s_override.Value;

                try
                {
                    var env = Environment.GetEnvironmentVariable("NATURALCOMMANDS_TEST_MODE");
                    if (!string.IsNullOrWhiteSpace(env))
                    {
                        if (env.Equals("1", StringComparison.OrdinalIgnoreCase) || env.Equals("true", StringComparison.OrdinalIgnoreCase))
                            return true;
                        if (env.Equals("0", StringComparison.OrdinalIgnoreCase) || env.Equals("false", StringComparison.OrdinalIgnoreCase))
                            return false;
                    }

                    // Backwards-compat: older tests used this env var
                    var old = Environment.GetEnvironmentVariable("NATURALCOMMANDS_TEST_NO_LAUNCH");
                    if (!string.IsNullOrWhiteSpace(old)) return true;
                }
                catch { }

                try
                {
                    // Allow settings.json override via AppSettings when available
                    return NaturalCommands.Models.AppSettings.Instance.Behavior.TestMode;
                }
                catch { }

                return false;
            }
        }

        /// <summary>
        /// Tests can call this to explicitly override test mode for the running process.
        /// </summary>
        public static void SetOverride(bool enabled) => s_override = enabled;

        public static void ClearOverride() => s_override = null;
    }
}
