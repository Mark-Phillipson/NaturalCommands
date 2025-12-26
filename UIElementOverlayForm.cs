using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using NaturalCommands.Helpers;

namespace NaturalCommands
{
    /// <summary>
    /// A transparent overlay form that displays letter labels on clickable UI elements.
    /// Used for the "show letters" voice navigation feature.
    /// </summary>
    public class UIElementOverlayForm : Form
    {
        private readonly Dictionary<string, UIAutomationHelper.ClickableElement> _labelToElementMap;
        private static UIElementOverlayForm? _currentInstance;
        private readonly Font _labelFont;
        private readonly Brush _labelBackgroundBrush;
        private readonly Brush _labelTextBrush;
        private readonly Pen _labelBorderPen;

        public UIElementOverlayForm(List<UIAutomationHelper.ClickableElement> elements)
        {
            _labelToElementMap = new Dictionary<string, UIAutomationHelper.ClickableElement>(
                StringComparer.OrdinalIgnoreCase);

            // Assign labels to elements
            var labels = UIAutomationHelper.GenerateLabels(elements.Count);
            for (int i = 0; i < elements.Count && i < labels.Count; i++)
            {
                elements[i].Label = labels[i];
                _labelToElementMap[labels[i]] = elements[i];
            }

            // Configure form appearance
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            BackColor = Color.Lime; // Use a bright color for transparency key
            TransparencyKey = Color.Lime;
            StartPosition = FormStartPosition.Manual;
            
            // Cover all screens
            var bounds = Screen.AllScreens.Aggregate(Rectangle.Empty, 
                (current, screen) => Rectangle.Union(current, screen.Bounds));
            Bounds = bounds;

            // Set up double buffering to reduce flicker
            DoubleBuffered = true;
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint | 
                     ControlStyles.OptimizedDoubleBuffer, true);

            // Initialize fonts and brushes
            _labelFont = new Font("Segoe UI", 10, FontStyle.Bold);
            _labelBackgroundBrush = new SolidBrush(Color.FromArgb(220, 255, 215, 0)); // Semi-transparent yellow
            _labelTextBrush = new SolidBrush(Color.Black);
            _labelBorderPen = new Pen(Color.FromArgb(255, 0, 0, 0), 3); // General border now 3px

            // Handle keyboard input for label selection
            KeyPreview = true;
            KeyPress += OnKeyPress;
            KeyDown += OnKeyDown;

            // Ensure overlay grabs focus and stays topmost when shown so it receives keyboard input
            Shown += (s, e) => {
                try
                {
                    this.Activate();
                    this.Focus();
                    this.BringToFront();
                    // Ensure topmost using SetWindowPos (HWND_TOPMOST = -1, SWP_NOMOVE|SWP_NOSIZE = 0x0001|0x0002)
                    Win32ApiHelper.SetWindowPos(this.Handle, new IntPtr(-1), 0, 0, 0, 0, 0x0001 | 0x0002);
                    Win32ApiHelper.SetForegroundWindow(this.Handle);
                    Logger.LogDebug($"Overlay Shown: activated and requested foreground (hwnd={this.Handle}).");
                }
                catch (Exception ex)
                {
                    Logger.LogError($"Error activating overlay: {ex.Message}");
                }
            };

            Logger.LogDebug($"Overlay created with {elements.Count} labeled elements.");
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            var graphics = e.Graphics;
            graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            // Draw labels for each element
            // Get virtual screen origin for correct label placement
            int virtualScreenX = SystemInformation.VirtualScreen.Left;
            int virtualScreenY = SystemInformation.VirtualScreen.Top;

            foreach (var kvp in _labelToElementMap)
            {
                var label = kvp.Key;
                var element = kvp.Value;

                // Position label at top-left corner of element, adjusted for virtual screen origin
                var labelSize = graphics.MeasureString(label.ToUpper(), _labelFont);
                float x = element.Bounds.Left - virtualScreenX;
                float y = element.Bounds.Top - virtualScreenY;
                var labelRect = new RectangleF(
                    x,
                    y,
                    labelSize.Width + 8,
                    labelSize.Height + 4);

                // Ensure label is visible on screen - constrain to screen bounds
                if (labelRect.X < 0)
                    labelRect.X = 0;
                if (labelRect.Y < 0)
                    labelRect.Y = 0;
                if (labelRect.Right > Bounds.Width)
                    labelRect.X = Bounds.Width - labelRect.Width;
                if (labelRect.Bottom > Bounds.Height)
                    labelRect.Y = Bounds.Height - labelRect.Height;

                // Draw label background
                graphics.FillRectangle(_labelBackgroundBrush, labelRect);
                // Draw a blue border for text boxes
                if (element.ControlType == "TextBox")
                {
                    using (var bluePen = new Pen(Color.Blue, 3)) // Blue border 3px
                        graphics.DrawRectangle(bluePen, Rectangle.Round(labelRect));
                }
                else
                {
                    graphics.DrawRectangle(_labelBorderPen, Rectangle.Round(labelRect));
                }

                // Draw label text
                var textPoint = new PointF(labelRect.X + 4, labelRect.Y + 2);
                graphics.DrawString(label.ToUpper(), _labelFont, _labelTextBrush, textPoint);
            }

            // Draw instruction text at bottom of screen
            var instruction = "Type letters to click elements. Press ESC to cancel.";
            var instructionSize = graphics.MeasureString(instruction, _labelFont);
            var instructionRect = new RectangleF(
                (Bounds.Width - instructionSize.Width) / 2,
                Bounds.Height - instructionSize.Height - 20,
                instructionSize.Width + 16,
                instructionSize.Height + 8);
            
            graphics.FillRectangle(new SolidBrush(Color.FromArgb(200, 0, 0, 0)), instructionRect);
            graphics.DrawString(instruction, _labelFont, Brushes.White, 
                instructionRect.X + 8, instructionRect.Y + 4);
        }

        private string _typedLabel = "";

        private void OnKeyPress(object? sender, KeyPressEventArgs e)
        {
            if (char.IsLetter(e.KeyChar))
            {
                _typedLabel += char.ToLower(e.KeyChar);
                Logger.LogDebug($"Typed label so far: {_typedLabel}");

                // Check if we have a match
                if (_labelToElementMap.TryGetValue(_typedLabel, out var element))
                {
                    Logger.LogDebug($"Found matching element for label: {_typedLabel}");
                    ClickElement(element);
                    e.Handled = true;
                }
                else if (_typedLabel.Length == 2)
                {
                    // Two characters typed but no match - reset
                    Logger.LogDebug($"No match for label: {_typedLabel}, resetting");
                    _typedLabel = "";
                }
                e.Handled = true;
            }
        }

        private void OnKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                Logger.LogDebug("ESC pressed, closing overlay.");
                Close();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Back && _typedLabel.Length > 0)
            {
                _typedLabel = _typedLabel.Substring(0, _typedLabel.Length - 1);
                Logger.LogDebug($"Backspace pressed, typed label: {_typedLabel}");
                e.Handled = true;
            }
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetCursorPos(int X, int Y);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, int dwExtraInfo);

        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;

        private void ClickElement(UIAutomationHelper.ClickableElement element)
        {
            try
            {
                Logger.LogDebug($"Clicking/focusing element: {element.Name} ({element.ControlType}) at {element.Bounds}");
                bool success = false;
                if (element.ControlType == "TextBox")
                {
                    try
                    {
                        element.Element.SetFocus();
                        success = true;
                        Logger.LogDebug("TextBox focused via SetFocus().");
                    }
                    catch (Exception ex)
                    {
                        Logger.LogError($"Failed to focus TextBox: {ex.Message}");
                    }
                    // Close overlay after focusing text box
                    Close();
                    return;
                }
                else
                {
                    // First try UI Automation patterns (Invoke, Toggle, SelectionItem)
                    success = UIAutomationHelper.ClickElement(element.Element);
                    if (success)
                    {
                        Logger.LogDebug("Element activated via UI Automation patterns.");
                        // Close the overlay and return
                        Close();
                        return;
                    }

                    // If UI Automation failed, fallback to mouse click at element center.
                    // We must close the overlay first so the click reaches the underlying window.
                    var centerX = element.Bounds.Left + (element.Bounds.Width / 2);
                    var centerY = element.Bounds.Top + (element.Bounds.Height / 2);

                    Logger.LogDebug($"UI Automation didn't activate element; falling back to mouse click at ({centerX},{centerY}).");

                    // Close overlay and give the system a brief moment to remove it
                    Close();
                    System.Windows.Forms.Application.DoEvents();
                    System.Threading.Thread.Sleep(90);

                    try
                    {
                        // Save current cursor position
                        var prevPos = System.Windows.Forms.Cursor.Position;
                        // Move cursor and click
                        SetCursorPos(centerX, centerY);
                        mouse_event(MOUSEEVENTF_LEFTDOWN, centerX, centerY, 0, 0);
                        mouse_event(MOUSEEVENTF_LEFTUP, centerX, centerY, 0, 0);
                        Logger.LogDebug("Performed mouse click fallback at element center.");

                        // Restore cursor after a short delay
                        System.Threading.Thread.Sleep(60);
                        SetCursorPos(prevPos.X, prevPos.Y);
                    }
                    catch (Exception ex)
                    {
                        Logger.LogError($"Mouse click fallback failed: {ex.Message}");
                    }
                }
                // Close the overlay after clicking/focusing (if not already closed)
                try { Close(); } catch { }
            }
            catch (Exception ex)
            {
                Logger.LogError($"Error clicking/focusing element: {ex.Message}");
                Close();
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _labelFont?.Dispose();
                _labelBackgroundBrush?.Dispose();
                _labelTextBrush?.Dispose();
                _labelBorderPen?.Dispose();
            }
            base.Dispose(disposing);
        }

        /// <summary>
        /// Shows the overlay with labeled UI elements. Only one instance can be active at a time.
        /// </summary>
        public static void ShowOverlay(bool scopeToActiveWindow = true)
        {
            try
            {
                // Close existing overlay if any
                CloseOverlay();

                var elements = UIAutomationHelper.EnumerateClickableElements(scopeToActiveWindow);
                
                if (elements.Count == 0)
                {
                    Logger.LogError("No clickable elements found.");
                    MessageBox.Show("No clickable elements found in the current window.", 
                        "Show Letters", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                _currentInstance = new UIElementOverlayForm(elements);
                
                // Use ShowDialog to keep the form alive and responsive
                // This blocks until the form is closed (via ESC or clicking an element)
                Logger.LogDebug($"Overlay shown with {elements.Count} elements.");
                _currentInstance.ShowDialog();
                
                // Clean up after dialog closes
                _currentInstance.Dispose();
                _currentInstance = null;
            }
            catch (Exception ex)
            {
                Logger.LogError($"Error showing overlay: {ex.Message}");
                MessageBox.Show($"Error showing overlay: {ex.Message}", 
                    "Show Letters Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Show overlay for a specific window handle (e.g., the Taskbar). This enumerates elements rooted at the hwnd.
        /// </summary>
        public static void ShowOverlayForWindow(IntPtr hwnd)
        {
            try
            {
                CloseOverlay();
                var elements = UIAutomationHelper.EnumerateClickableElementsForWindow(hwnd);
                if (elements.Count == 0)
                {
                    Logger.LogError($"No clickable elements found for hwnd {hwnd}.");
                    MessageBox.Show("No clickable elements found in the target window.", 
                        "Show Letters", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                _currentInstance = new UIElementOverlayForm(elements);
                Logger.LogDebug($"Overlay shown for hwnd {hwnd} with {elements.Count} elements.");
                _currentInstance.ShowDialog();
                _currentInstance.Dispose();
                _currentInstance = null;
            }
            catch (Exception ex)
            {
                Logger.LogError($"Error showing overlay for hwnd {hwnd}: {ex.Message}");
                MessageBox.Show($"Error showing overlay: {ex.Message}", 
                    "Show Letters Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Closes the current overlay if one is active.
        /// </summary>
        public static void CloseOverlay()
        {
            if (_currentInstance != null && !_currentInstance.IsDisposed)
            {
                try
                {
                    _currentInstance.Close();
                    _currentInstance.Dispose();
                }
                catch { }
                _currentInstance = null;
            }
        }
    }
}
