using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using UIAutomationClient;

namespace NaturalCommands.Helpers
{
    /// <summary>
    /// Helper class for enumerating UI elements using Windows UI Automation API.
    /// Used to find clickable elements for the "show letters" feature.
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
            public IUIAutomationElement Element { get; set; } = null!;
            public string Name { get; set; } = "";
            public string ControlType { get; set; } = "";
        }

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
                var automation = new CUIAutomation();
                IUIAutomationElement rootElement;

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

                // Create condition for clickable elements (buttons, links, menu items, etc.)
                var condition = CreateClickableCondition(automation);
                var walker = automation.CreateTreeWalker(condition);

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

        private static IUIAutomationCondition CreateClickableCondition(IUIAutomation automation)
        {
            // Create conditions for various clickable control types
            var buttonCondition = automation.CreatePropertyCondition(
                UIA_PropertyIds.UIA_ControlTypePropertyId,
                UIA_ControlTypeIds.UIA_ButtonControlTypeId);

            var linkCondition = automation.CreatePropertyCondition(
                UIA_PropertyIds.UIA_ControlTypePropertyId,
                UIA_ControlTypeIds.UIA_HyperlinkControlTypeId);

            var menuItemCondition = automation.CreatePropertyCondition(
                UIA_PropertyIds.UIA_ControlTypePropertyId,
                UIA_ControlTypeIds.UIA_MenuItemControlTypeId);

            var checkBoxCondition = automation.CreatePropertyCondition(
                UIA_PropertyIds.UIA_ControlTypePropertyId,
                UIA_ControlTypeIds.UIA_CheckBoxControlTypeId);

            var radioButtonCondition = automation.CreatePropertyCondition(
                UIA_PropertyIds.UIA_ControlTypePropertyId,
                UIA_ControlTypeIds.UIA_RadioButtonControlTypeId);

            var tabItemCondition = automation.CreatePropertyCondition(
                UIA_PropertyIds.UIA_ControlTypePropertyId,
                UIA_ControlTypeIds.UIA_TabItemControlTypeId);

            var comboBoxCondition = automation.CreatePropertyCondition(
                UIA_PropertyIds.UIA_ControlTypePropertyId,
                UIA_ControlTypeIds.UIA_ComboBoxControlTypeId);

            // Combine all conditions with OR
            var conditions = new IUIAutomationCondition[]
            {
                buttonCondition, linkCondition, menuItemCondition, 
                checkBoxCondition, radioButtonCondition, tabItemCondition, comboBoxCondition
            };

            return automation.CreateOrConditionFromArray(conditions);
        }

        private static void WalkElements(IUIAutomationElement element, IUIAutomation automation, 
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
                        var name = element.CurrentName ?? "";
                        var controlType = GetControlTypeName(element.CurrentControlType);

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
                var walker = automation.ControlViewWalker;
                var child = walker.GetFirstChildElement(element);
                while (child != null && clickableElements.Count < maxElements)
                {
                    WalkElements(child, automation, clickableElements, depth + 1, maxElements);
                    child = walker.GetNextSiblingElement(child);
                }
            }
            catch (Exception ex)
            {
                // Some elements may throw exceptions when accessed, just skip them
                Logger.LogDebug($"Skipped element due to error: {ex.Message}");
            }
        }

        private static bool IsElementClickable(IUIAutomationElement element)
        {
            try
            {
                // Check if element is enabled and on-screen
                if (element.CurrentIsEnabled == 0 || element.CurrentIsOffscreen != 0)
                    return false;

                // Check if element supports Invoke or Toggle patterns (clickable)
                object invokePattern = null;
                object togglePattern = null;
                try
                {
                    invokePattern = element.GetCurrentPattern(UIA_PatternIds.UIA_InvokePatternId);
                }
                catch { }

                try
                {
                    togglePattern = element.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId);
                }
                catch { }

                return invokePattern != null || togglePattern != null;
            }
            catch
            {
                return false;
            }
        }

        private static Rectangle GetElementBounds(IUIAutomationElement element)
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
                UIA_ControlTypeIds.UIA_ButtonControlTypeId => "Button",
                UIA_ControlTypeIds.UIA_HyperlinkControlTypeId => "Link",
                UIA_ControlTypeIds.UIA_MenuItemControlTypeId => "MenuItem",
                UIA_ControlTypeIds.UIA_CheckBoxControlTypeId => "CheckBox",
                UIA_ControlTypeIds.UIA_RadioButtonControlTypeId => "RadioButton",
                UIA_ControlTypeIds.UIA_TabItemControlTypeId => "TabItem",
                UIA_ControlTypeIds.UIA_ComboBoxControlTypeId => "ComboBox",
                _ => "Unknown"
            };
        }

        /// <summary>
        /// Clicks on a UI element using UI Automation.
        /// </summary>
        public static bool ClickElement(IUIAutomationElement element)
        {
            try
            {
                // Try Invoke pattern first (most common for clickable elements)
                try
                {
                    var invokePattern = element.GetCurrentPattern(UIA_PatternIds.UIA_InvokePatternId) 
                        as IUIAutomationInvokePattern;
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
                    var togglePattern = element.GetCurrentPattern(UIA_PatternIds.UIA_TogglePatternId) 
                        as IUIAutomationTogglePattern;
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
