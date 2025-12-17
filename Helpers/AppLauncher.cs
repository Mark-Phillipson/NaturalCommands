using System;
using System.IO;
namespace NaturalCommands.Helpers
{
    // Handles application launching and focusing
    public class AppLauncher
    {
        // TODO: Move app launching/focusing methods here from NaturalLanguageInterpreter
        public static string Launch(NaturalCommands.LaunchAppAction app)
        {
            try
            {
                string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "bin", "app.log");
                logPath = Path.GetFullPath(logPath);
                try { NaturalCommands.Helpers.Logger.LogDebug($"AppLauncher.Launch: Requested launch '{app?.AppExe}'"); } catch { }

                if (app == null || string.IsNullOrWhiteSpace(app.AppExe))
                {
                    try { NaturalCommands.Helpers.Logger.LogError("AppLauncher.Launch: No app specified."); } catch { }
                    return "No application specified to launch.";
                }

                var psi = new System.Diagnostics.ProcessStartInfo(app.AppExe)
                {
                    UseShellExecute = true
                };

                try
                {
                    System.Diagnostics.Process.Start(psi);
                    try { NaturalCommands.Helpers.Logger.LogInfo($"AppLauncher.Launch: Started '{app.AppExe}' successfully."); } catch { }
                    return $"Launched: {app.AppExe}";
                }
                catch (Exception ex)
                {
                    try { NaturalCommands.Helpers.Logger.LogError($"AppLauncher.Launch: Failed to start '{app.AppExe}': {ex.Message}"); } catch { }
                    // Try fallback: use explorer to open by verb
                    try
                    {
                        var psi2 = new System.Diagnostics.ProcessStartInfo("explorer.exe", app.AppExe) { UseShellExecute = true };
                        System.Diagnostics.Process.Start(psi2);
                        try { NaturalCommands.Helpers.Logger.LogInfo($"AppLauncher.Launch: Fallback explorer launched for '{app.AppExe}'."); } catch { }
                        return $"Launched via explorer: {app.AppExe}";
                    }
                    catch (Exception ex2)
                    {
                        try { NaturalCommands.Helpers.Logger.LogError($"AppLauncher.Launch: Explorer fallback failed for '{app.AppExe}': {ex2.Message}"); } catch { }
                        return $"Failed to launch '{app.AppExe}': {ex.Message}; fallback: {ex2.Message}";
                    }
                }
            }
            catch (Exception exOuter)
            {
                NaturalCommands.Helpers.Logger.LogError($"AppLauncher.Launch: Unexpected: {exOuter.Message}");
                return $"App launch failed: {exOuter.Message}";
            }
        }
    }
}
