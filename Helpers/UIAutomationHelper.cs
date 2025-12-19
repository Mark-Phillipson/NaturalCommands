using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Automation;

namespace NaturalCommands.Helpers
{
    /// <summary>
    /// Helper class for enumerating UI elements using Windows UI Automation API.
    /// Used to find clickable elements for the "show letters" feature.
    /// Uses managed System.Windows.Automation API.
    /// </summary>
    public static class UIAutomationHelper
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        /// <summary>
        /// Represents a clickable UI element with its position and label.
        /// </summary>
        public class ClickableElement
        {
            public Rectangle Bounds { get; set; }
            public string Label { get; set; } = "";
            public AutomationElement Element { get; set; } = null!;
            public string Name { get; set; } = "";
            public string ControlType { get; set; } = "";
        }

        /// <summary>
        /// Generates labels using the Talon alphabet.
        /// Uses single letters (a-z) for 26 or fewer elements for efficiency.
        /// Uses two-letter combinations (aa, ab, ac, ...) for more than 26 elements to avoid instant activation.
        /// </summary>
        public static List<string> GenerateLabels(int count)
        {
            var labels = new List<string>();
            var alphabet = "abcdefghijklmnopqrstuvwxyz";

            // Use single letters if we have 26 or fewer elements
            if (count <= 26)
            {
                foreach (var letter in alphabet)
                {
                    if (labels.Count >= count) break;
                    labels.Add(letter.ToString());
                }
            }
            else
            {
                // Use two-letter combinations for more than 26 elements
                foreach (var first in alphabet)
                {
                    foreach (var second in alphabet)
                    {
                        if (labels.Count >= count) break;
                        labels.Add($"{first}{second}");
                    }
                    if (labels.Count >= count) break;
                }
            }

            return labels.Take(count).ToList();
        }

        /// <summary>
        /// Enumerates clickable UI elements on the screen or within the active window.
        /// </summary>
        /// <param name="scopeToActiveWindow">If true, only search within the active window. If false, search entire screen.</param>
        public static List<ClickableElement> EnumerateClickableElements(bool scopeToActiveWindow = true)
        {
            var clickableElements = new List<ClickableElement>();

            try
            {
                AutomationElement rootElement;

                if (scopeToActiveWindow)
                {
                    IntPtr hwnd = GetForegroundWindow();
                    if (hwnd == IntPtr.Zero)
                    {
                        Logger.LogError("Could not get foreground window handle.");
                        return clickableElements;
                    }
                    rootElement = AutomationElement.FromHandle(hwnd);
                }
                else
                {
                    rootElement = AutomationElement.RootElement;
                }

                if (rootElement == null)
                {
                    Logger.LogError("Could not get root UI Automation element.");
                    return clickableElements;
                }

                // Walk the tree and collect clickable elements
                WalkElements(rootElement, clickableElements, 0, 500); // Limit to 500 elements

                Logger.LogDebug($"Found {clickableElements.Count} clickable elements.");
            }
            catch (Exception ex)
            {
                Logger.LogError($"Error enumerating clickable elements: {ex.Message}");
            }

            return clickableElements;
        }

        private static void WalkElements(AutomationElement element, 
            List<ClickableElement> clickableElements, int depth, int maxElements)
        {
            if (element == null || clickableElements.Count >= maxElements || depth > 20)
                return;

            try
            {
                // Check if element is clickable and visible
                if (IsElementClickable(element))
                {
                    var bounds = GetElementBounds(element);
                    if (bounds.Width > 0 && bounds.Height > 0)
                    {
                        string name = "";
                        try { name = element.Current.Name ?? ""; } catch { }
                        
                        var controlType = GetControlTypeName(element.Current.ControlType);

                        clickableElements.Add(new ClickableElement
                        {
                            Bounds = bounds,
                            Element = element,
                            Name = name,
                            ControlType = controlType
                        });
                    }
                }

                // Recursively walk children using TreeWalker
                try
                {
                    var walker = TreeWalker.ControlViewWalker;
                    var child = walker.GetFirstChild(element);
                    while (child != null && clickableElements.Count < maxElements)
                    {
                        WalkElements(child, clickableElements, depth + 1, maxElements);
                        try
                        {
                            child = walker.GetNextSibling(child);
                        }
                        catch
                        {
                            child = null;
                        }
                    }
                }
                catch { }
            }
            catch (Exception ex)
            {
                // Some elements may throw exceptions when accessed, just skip them
                Logger.LogDebug($"Skipped element due to error: {ex.Message}");
            }
        }

        private static bool IsElementClickable(AutomationElement element)
        {
            try
            {
                // Check if element is enabled and on-screen
                if (!element.Current.IsEnabled || element.Current.IsOffscreen)
                    return false;

                // Check if element supports Invoke or Toggle patterns (clickable)
                object? invokePattern = null;
                object? togglePattern = null;
                
                try
                {
                    invokePattern = element.GetCurrentPattern(InvokePattern.Pattern);
                }
                catch { }

                try
                {
                    togglePattern = element.GetCurrentPattern(TogglePattern.Pattern);
                }
                catch { }

                return invokePattern != null || togglePattern != null;
            }
            catch
            {
                return false;
            }
        }

        private static Rectangle GetElementBounds(AutomationElement element)
        {
            try
            {
                var rect = element.Current.BoundingRectangle;
                return new Rectangle(
                    (int)rect.Left,
                    (int)rect.Top,
                    (int)rect.Width,
                    (int)rect.Height
                );
            }
            catch
            {
                return Rectangle.Empty;
            }
        }

        private static string GetControlTypeName(ControlType controlType)
        {
            if (controlType == ControlType.Button) return "Button";
            if (controlType == ControlType.Hyperlink) return "Link";
            if (controlType == ControlType.MenuItem) return "MenuItem";
            if (controlType == ControlType.CheckBox) return "CheckBox";
            if (controlType == ControlType.RadioButton) return "RadioButton";
            if (controlType == ControlType.TabItem) return "TabItem";
            if (controlType == ControlType.ComboBox) return "ComboBox";
            return "Unknown";
        }

        /// <summary>
        /// Clicks on a UI element using UI Automation.
        /// </summary>
        public static bool ClickElement(AutomationElement element)
        {
            try
            {
                // Try Invoke pattern first (most common for clickable elements)
                try
                {
                    var invokePattern = element.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                    if (invokePattern != null)
                    {
                        invokePattern.Invoke();
                        return true;
                    }
                }
                catch { }

                // Try Toggle pattern for checkboxes and radio buttons
                try
                {
                    var togglePattern = element.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                    if (togglePattern != null)
                    {
                        togglePattern.Toggle();
                        return true;
                    }
                }
                catch { }

                Logger.LogError("Element does not support Invoke or Toggle patterns.");
                return false;
            }
            catch (Exception ex)
            {
                Logger.LogError($"Error clicking element: {ex.Message}");
                return false;
            }
        }
    }
}
