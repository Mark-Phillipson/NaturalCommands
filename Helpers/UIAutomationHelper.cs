using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;

namespace NaturalCommands.Helpers
{
    /// <summary>
    /// Helper class for enumerating UI elements using Windows UI Automation API.
    /// Used to find clickable elements for the "show letters" feature.
    /// Uses dynamic COM to avoid compile-time dependencies.
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
            public dynamic Element { get; set; } = null!;
            public string Name { get; set; } = "";
            public string ControlType { get; set; } = "";
        }

        // UI Automation control type IDs
        private const int UIA_ButtonControlTypeId = 50000;
        private const int UIA_HyperlinkControlTypeId = 50005;
        private const int UIA_MenuItemControlTypeId = 50011;
        private const int UIA_CheckBoxControlTypeId = 50002;
        private const int UIA_RadioButtonControlTypeId = 50013;
        private const int UIA_TabItemControlTypeId = 50019;
        private const int UIA_ComboBoxControlTypeId = 50003;

        // UI Automation property IDs
        private const int UIA_ControlTypePropertyId = 30003;

        // UI Automation pattern IDs
        private const int UIA_InvokePatternId = 10000;
        private const int UIA_TogglePatternId = 10015;

        /// <summary>
        /// Generates two-letter labels using the Talon alphabet (a, b, c, ..., z, then aa, ab, ...).
        /// </summary>
        public static List<string> GenerateLabels(int count)
        {
            var labels = new List<string>();
            var alphabet = "abcdefghijklmnopqrstuvwxyz";

            // Single letters first
            foreach (var letter in alphabet)
            {
                if (labels.Count >= count) break;
                labels.Add(letter.ToString());
            }

            // Two-letter combinations
            foreach (var first in alphabet)
            {
                foreach (var second in alphabet)
                {
                    if (labels.Count >= count) break;
                    labels.Add($"{first}{second}");
                }
                if (labels.Count >= count) break;
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
                // Create UI Automation COM object dynamically
                Type? automationType = Type.GetTypeFromProgID("UIAutomationClient.CUIAutomation");
                if (automationType == null)
                {
                    Logger.LogError("Could not get UIAutomationClient COM type.");
                    return clickableElements;
                }

                dynamic automation = Activator.CreateInstance(automationType)!;
                dynamic rootElement;

                if (scopeToActiveWindow)
                {
                    IntPtr hwnd = GetForegroundWindow();
                    if (hwnd == IntPtr.Zero)
                    {
                        Logger.LogError("Could not get foreground window handle.");
                        return clickableElements;
                    }
                    rootElement = automation.ElementFromHandle(hwnd);
                }
                else
                {
                    rootElement = automation.GetRootElement();
                }

                if (rootElement == null)
                {
                    Logger.LogError("Could not get root UI Automation element.");
                    return clickableElements;
                }

                // Walk the tree and collect clickable elements
                WalkElements(rootElement, automation, clickableElements, 0, 500); // Limit to 500 elements

                Logger.LogDebug($"Found {clickableElements.Count} clickable elements.");
            }
            catch (Exception ex)
            {
                Logger.LogError($"Error enumerating clickable elements: {ex.Message}");
            }

            return clickableElements;
        }

        private static void WalkElements(dynamic element, dynamic automation, 
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
                        try { name = element.CurrentName ?? ""; } catch { }
                        
                        int controlTypeId = 0;
                        try { controlTypeId = element.CurrentControlType; } catch { }
                        var controlType = GetControlTypeName(controlTypeId);

                        clickableElements.Add(new ClickableElement
                        {
                            Bounds = bounds,
                            Element = element,
                            Name = name,
                            ControlType = controlType
                        });
                    }
                }

                // Recursively walk children
                try
                {
                    var walker = automation.ControlViewWalker;
                    var child = walker.GetFirstChildElement(element);
                    while (child != null && clickableElements.Count < maxElements)
                    {
                        WalkElements(child, automation, clickableElements, depth + 1, maxElements);
                        try
                        {
                            child = walker.GetNextSiblingElement(child);
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

        private static bool IsElementClickable(dynamic element)
        {
            try
            {
                // Check if element is enabled and on-screen
                int isEnabled = 0;
                int isOffscreen = 1;
                try { isEnabled = element.CurrentIsEnabled; } catch { }
                try { isOffscreen = element.CurrentIsOffscreen; } catch { }
                
                if (isEnabled == 0 || isOffscreen != 0)
                    return false;

                // Check if element supports Invoke or Toggle patterns (clickable)
                object? invokePattern = null;
                object? togglePattern = null;
                try
                {
                    invokePattern = element.GetCurrentPattern(UIA_InvokePatternId);
                }
                catch { }

                try
                {
                    togglePattern = element.GetCurrentPattern(UIA_TogglePatternId);
                }
                catch { }

                return invokePattern != null || togglePattern != null;
            }
            catch
            {
                return false;
            }
        }

        private static Rectangle GetElementBounds(dynamic element)
        {
            try
            {
                var rect = element.CurrentBoundingRectangle;
                return new Rectangle(
                    (int)rect.left,
                    (int)rect.top,
                    (int)(rect.right - rect.left),
                    (int)(rect.bottom - rect.top)
                );
            }
            catch
            {
                return Rectangle.Empty;
            }
        }

        private static string GetControlTypeName(int controlTypeId)
        {
            return controlTypeId switch
            {
                UIA_ButtonControlTypeId => "Button",
                UIA_HyperlinkControlTypeId => "Link",
                UIA_MenuItemControlTypeId => "MenuItem",
                UIA_CheckBoxControlTypeId => "CheckBox",
                UIA_RadioButtonControlTypeId => "RadioButton",
                UIA_TabItemControlTypeId => "TabItem",
                UIA_ComboBoxControlTypeId => "ComboBox",
                _ => "Unknown"
            };
        }

        /// <summary>
        /// Clicks on a UI element using UI Automation.
        /// </summary>
        public static bool ClickElement(dynamic element)
        {
            try
            {
                // Try Invoke pattern first (most common for clickable elements)
                try
                {
                    var invokePattern = element.GetCurrentPattern(UIA_InvokePatternId);
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
                    var togglePattern = element.GetCurrentPattern(UIA_TogglePatternId);
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
