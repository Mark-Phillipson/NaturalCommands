using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;

namespace NaturalCommands
{
    /// <summary>
    /// Generates the application icon dynamically.
    /// </summary>
    public static class AppIconGenerator
    {
        public enum IconState
        {
            Normal,
            AutoClickActive,
            ListenerOk,
            ListenerWarning,
            ListenerError,
            Unknown
        }
        /// <summary>
        /// Creates a custom icon for the system tray with a microphone design.
        /// </summary>
        public static Icon CreateAppIcon(bool isAutoClickActive = false)
        {
            return CreateAppIcon(isAutoClickActive ? IconState.AutoClickActive : IconState.Normal);
        }

        public static Icon CreateAppIcon(IconState state)
        {
            // Create a bitmap for the icon (32x32 for high quality)
            using (Bitmap bmp = new Bitmap(32, 32, PixelFormat.Format32bppArgb))
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                // Map icon state to colors
                Color primaryColor;
                Color accentColor;
                Color secondaryColor = Color.White;
                switch (state)
                {
                    case IconState.AutoClickActive:
                    case IconState.ListenerOk:
                        primaryColor = Color.FromArgb(16, 185, 129); // green
                        accentColor = Color.FromArgb(5, 150, 105);
                        break;
                    case IconState.ListenerWarning:
                        primaryColor = Color.FromArgb(245, 158, 11); // amber
                        accentColor = Color.FromArgb(202, 102, 0);
                        break;
                    case IconState.ListenerError:
                        primaryColor = Color.FromArgb(239, 68, 68); // red
                        accentColor = Color.FromArgb(180, 30, 30);
                        break;
                    case IconState.Unknown:
                        primaryColor = Color.FromArgb(120, 120, 120); // gray
                        accentColor = Color.FromArgb(90, 90, 90);
                        break;
                    case IconState.Normal:
                    default:
                        primaryColor = Color.FromArgb(0, 120, 215); // blue
                        accentColor = Color.FromArgb(0, 90, 158);
                        break;
                }

                // Draw microphone body
                using (var brush = new SolidBrush(primaryColor))
                {
                    // Microphone capsule (rounded rectangle)
                    g.FillEllipse(brush, 10, 4, 12, 16);
                }

                // Draw microphone stand/handle
                using (var pen = new Pen(primaryColor, 2.5f))
                {
                    pen.StartCap = LineCap.Round;
                    pen.EndCap = LineCap.Round;
                    
                    // Vertical line down
                    g.DrawLine(pen, 16, 20, 16, 26);
                    
                    // Horizontal base
                    g.DrawLine(pen, 12, 26, 20, 26);
                    
                    // Arc from bottom of mic to stand
                    g.DrawArc(pen, 11, 18, 10, 8, 0, 180);
                }

                // Add sound wave indicator (small accent)
                using (var pen = new Pen(accentColor, 1.5f))
                {
                    pen.StartCap = LineCap.Round;
                    pen.EndCap = LineCap.Round;
                    
                    // Small sound waves on the left
                    g.DrawArc(pen, 2, 8, 6, 8, -30, 60);
                    g.DrawArc(pen, 1, 6, 8, 12, -30, 60);
                    
                    // Small sound waves on the right
                    g.DrawArc(pen, 24, 8, 6, 8, 150, 60);
                    g.DrawArc(pen, 23, 6, 8, 12, 150, 60);
                }

                // Convert bitmap to icon
                IntPtr hIcon = bmp.GetHicon();
                Icon icon = Icon.FromHandle(hIcon);
                
                // Clone the icon so we can properly dispose of resources
                Icon clonedIcon = (Icon)icon.Clone();
                
                // Clean up
                DestroyIcon(hIcon);
                icon.Dispose();
                
                return clonedIcon;
            }
        }

        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private extern static bool DestroyIcon(IntPtr handle);
    }
}
