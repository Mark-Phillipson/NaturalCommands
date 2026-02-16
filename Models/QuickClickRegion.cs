using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace NaturalCommands.Models
{
    /// <summary>
    /// Represents a named set of clickable regions for a specific process/window.
    /// Profiles may be scoped to a specific monitor resolution (MonitorWidth/MonitorHeight)
    /// to keep pixel coordinates accurate across displays.
    /// </summary>
    public class QuickClickProfile
    {
        /// <summary>Process name (e.g. "chrome") - matched exactly.</summary>
        public string ProcessName { get; set; } = string.Empty;

        /// <summary>Optional window-title substring or pattern. Null/empty = any title.</summary>
        public string? WindowTitlePattern { get; set; }

        /// <summary>Optional recorded monitor width in pixels. Null = apply to any monitor.</summary>
        public int? MonitorWidth { get; set; }

        /// <summary>Optional recorded monitor height in pixels. Null = apply to any monitor.</summary>
        public int? MonitorHeight { get; set; }

        /// <summary>Collection of regions belonging to this profile.</summary>
        public List<QuickClickRegion> Regions { get; set; } = new List<QuickClickRegion>();
    }

    /// <summary>
    /// A single clickable region within a QuickClickProfile.
    /// Coordinates are stored as pixel offsets from the target window's client-area top-left.
    /// </summary>
    public class QuickClickRegion
    {
        /// <summary>Display name (used for voice matching).</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>X offset in pixels from the window client area's left.</summary>
        public int X { get; set; }

        /// <summary>Y offset in pixels from the window client area's top.</summary>
        public int Y { get; set; }

        /// <summary>Region width in pixels.</summary>
        public int Width { get; set; } = 100;

        /// <summary>Region height in pixels.</summary>
        public int Height { get; set; } = 100;

        /// <summary>Default click type for the editor / when explicitly requested.</summary>
        public QuickClickClickType ClickType { get; set; } = QuickClickClickType.Left;

        /// <summary>
        /// Optional explicit voice command for this region. If null/empty the region's
        /// <see cref="Name"/> lower-cased will be used by the interpreter.
        /// </summary>
        public string? VoiceCommand { get; set; }

        /// <summary>Helper: returns the effective voice command (VoiceCommand or Name.ToLowerInvariant()).</summary>
        [JsonIgnore]
        public string EffectiveVoiceCommand =>
            !string.IsNullOrWhiteSpace(VoiceCommand) ? VoiceCommand! : Name.ToLowerInvariant();
    }

    /// <summary>
    /// Supported click types for Quick Click regions.
    /// Serialized as text (string) in JSON.
    /// </summary>
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public enum QuickClickClickType
    {
        Left,
        Double,
        Right,
        Middle
    }
}
