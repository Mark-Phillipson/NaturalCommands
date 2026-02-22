using System;
using System.Collections.Generic;

namespace NaturalCommands
{
    // Action type for Visual Studio command execution
    public abstract record ActionBase;
    public record RunMultipleActionsAction(string Name, List<ActionBase> Actions, bool ContinueOnError = true, int DelayMsBetween = 250) : ActionBase;
    public record MoveWindowAction(string Target, string Monitor, string? Position, int? WidthPercent, int? HeightPercent) : ActionBase;
    // Default TimeoutMs = 0 disables auto-submit while allowing auto-start
    // (the form only auto-submits when TimeoutMs > 0).
    public record OpenVoiceDictationFormAction(int TimeoutMs = 0) : ActionBase;
    public record OpenSettingsAction : ActionBase;
    public record CloseTabAction : ActionBase { }
    public record SetWindowAlwaysOnTopAction(string? Application) : ActionBase;
    public record ExecuteVSCommandAction(string CommandName, string? Arguments = null) : ActionBase;
    public record EmojiAction(string? Name, string EmojiText) : ActionBase;
    public record OpenFolderAction(string KnownFolder) : ActionBase;
    public record FocusWindowAction(string WindowTitleSubstring) : ActionBase;
    public record OpenWebsiteAction(string Url) : ActionBase;
    public record ShowHelpAction : ActionBase;
    public record LaunchAppAction(string AppExe) : ActionBase;
    public record SendKeysAction(string KeysText) : ActionBase;
    public record ShowLettersAction(bool ScopeToActiveWindow = true) : ActionBase;
    public record ShowTaskbarAction : ActionBase;
    public record ShowDesktopAction : ActionBase;
    public record StartMouseMoveAction(string Direction) : ActionBase;
    public record StopMouseMoveAction(bool PerformClick = false, bool IsRightClick = false) : ActionBase;
    public record AdjustMouseSpeedAction(string SpeedChange) : ActionBase;

    // Click a named quick-click region (ClickType optional; defaults to Left when omitted)
    // Uses the model enum `Models.QuickClickClickType` defined in Models/QuickClickRegion.cs.
    public record ClickQuickClickAction(string RegionName, Models.QuickClickClickType ClickType = Models.QuickClickClickType.Left) : ActionBase;

    public record ShowQuickClicksAction() : ActionBase;
    public record HideQuickClicksAction() : ActionBase;
    public record EditQuickClicksAction() : ActionBase;
    public record CreateQuickClickAction() : ActionBase;
    // Adjust auto-click delay by spoken commands such as "auto click faster" / "auto click slower"
    public record AdjustAutoClickDelayAction(string Change) : ActionBase;
    public record StartAutoClickAction(int DelayMs = 0) : ActionBase;
    public record StopAutoClickAction : ActionBase;
    // Windows Terminal specific shortcut action
    public record WindowsTerminalShortcutAction(string Shortcut, string CommandText) : ActionBase;
    // Run a terminal command (types command text and presses Enter)
    public record RunTerminalCommandAction(string Command, string Description) : ActionBase;
    // Windows Explorer specific shortcut action
    public record WindowsExplorerShortcutAction(string Shortcut, string CommandText) : ActionBase;
    // Execute a matched Talon voice command using configured bridge settings.
    public record RunTalonCommandAction(string OriginalText, string TalonCommand, string MatchSource = "catalog") : ActionBase;
    // Reload Talon command catalog from Talon user scripts.
    public record RefreshTalonCatalogAction : ActionBase;
    public record VisualIdentifyClickAction(string TargetPhrase) : ActionBase;
    public record VisualShowCandidatesAction : ActionBase;
    public record VisualChooseCandidateAction(int CandidateNumber) : ActionBase;
}
