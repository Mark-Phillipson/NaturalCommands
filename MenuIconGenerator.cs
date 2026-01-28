using System.Drawing;
using System.Drawing.Drawing2D;

namespace NaturalCommands
{
    /// <summary>
    /// Creates small bitmap icons used for tray menu items.
    /// </summary>
    public static class MenuIconGenerator
    {
        public static Image CreateMicrophoneImage(int size = 16)
        {
            var bmp = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                var primary = Color.FromArgb(0, 120, 215);
                using (var brush = new SolidBrush(primary))
                using (var pen = new Pen(primary, 1.5f))
                {
                    // Mic capsule
                    g.FillEllipse(brush, size * 0.35f, size * 0.15f, size * 0.3f, size * 0.45f);
                    // Stand
                    g.DrawLine(pen, size * 0.5f, size * 0.6f, size * 0.5f, size * 0.78f);
                    g.DrawLine(pen, size * 0.4f, size * 0.78f, size * 0.6f, size * 0.78f);
                }

                // Waves
                using (var pen = new Pen(Color.FromArgb(0, 90, 158), 1f))
                {
                    g.DrawArc(pen, size * 0.08f, size * 0.12f, size * 0.2f, size * 0.35f, -45, 90);
                    g.DrawArc(pen, size * 0.72f, size * 0.12f, size * 0.2f, size * 0.35f, 135, 90);
                }
            }

            return bmp;
        }

        public static Image CreateStopImage(int size = 16)
        {
            var bmp = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                var red = Color.FromArgb(220, 38, 38);
                using (var brush = new SolidBrush(red))
                using (var pen = new Pen(Color.White, 1.5f))
                {
                    // Rounded square
                    var rect = new RectangleF(size * 0.18f, size * 0.18f, size * 0.64f, size * 0.64f);
                    var path = new GraphicsPath();
                    float r = size * 0.12f;
                    path.AddArc(rect.X, rect.Y, r, r, 180, 90);
                    path.AddArc(rect.Right - r, rect.Y, r, r, 270, 90);
                    path.AddArc(rect.Right - r, rect.Bottom - r, r, r, 0, 90);
                    path.AddArc(rect.X, rect.Bottom - r, r, r, 90, 90);
                    path.CloseFigure();
                    g.FillPath(brush, path);
                }
            }

            return bmp;
        }

        public static Image CreateSettingsImage(int size = 16)
        {
            var bmp = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                using (var brush = new SolidBrush(Color.FromArgb(80, 80, 80)))
                using (var pen = new Pen(Color.FromArgb(80, 80, 80), 1.2f))
                {
                    // Central circle
                    float cx = size / 2f, cy = size / 2f;
                    float r = size * 0.18f;
                    g.FillEllipse(brush, cx - r, cy - r, r * 2, r * 2);

                    // 6 teeth around
                    for (int i = 0; i < 6; i++)
                    {
                        float angle = i * 60 * (float)(Math.PI / 180);
                        float tx = cx + (float)Math.Cos(angle) * size * 0.34f;
                        float ty = cy + (float)Math.Sin(angle) * size * 0.34f;
                        var rect = new RectangleF(tx - size * 0.07f, ty - size * 0.07f, size * 0.14f, size * 0.14f);
                        g.FillRectangle(brush, rect);
                    }
                }
            }

            return bmp;
        }

        public static Image CreateExitImage(int size = 16)
        {
            var bmp = new Bitmap(size, size);
            using (var g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);

                using (var pen = new Pen(Color.FromArgb(90, 90, 90), 1.8f))
                {
                    // Circle
                    g.DrawEllipse(pen, size * 0.12f, size * 0.12f, size * 0.76f, size * 0.76f);

                    // X
                    g.DrawLine(pen, size * 0.36f, size * 0.36f, size * 0.64f, size * 0.64f);
                    g.DrawLine(pen, size * 0.64f, size * 0.36f, size * 0.36f, size * 0.64f);
                }
            }

            return bmp;
        }
    }
}
