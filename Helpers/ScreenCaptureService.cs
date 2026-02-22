using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace NaturalCommands.Helpers
{
    public class ScreenCaptureResult
    {
        public Rectangle VirtualBounds { get; set; }
        public string ImageBase64Jpeg { get; set; } = string.Empty;
        public int Width { get; set; }
        public int Height { get; set; }
    }

    public static class ScreenCaptureService
    {
        public static ScreenCaptureResult CaptureAllScreensJpeg(int maxLongEdge = 1400, long quality = 55)
        {
            var bounds = Screen.AllScreens.Aggregate(Rectangle.Empty, (current, screen) => Rectangle.Union(current, screen.Bounds));
            using var bitmap = new Bitmap(bounds.Width, bounds.Height);
            using (var graphics = Graphics.FromImage(bitmap))
            {
                graphics.CopyFromScreen(bounds.Left, bounds.Top, 0, 0, bounds.Size, CopyPixelOperation.SourceCopy);
            }

            using var resized = ResizeToMaxLongEdge(bitmap, maxLongEdge);
            using var memory = new MemoryStream();
            SaveAsJpeg(resized, memory, quality);

            return new ScreenCaptureResult
            {
                VirtualBounds = bounds,
                Width = resized.Width,
                Height = resized.Height,
                ImageBase64Jpeg = Convert.ToBase64String(memory.ToArray())
            };
        }

        private static Bitmap ResizeToMaxLongEdge(Bitmap source, int maxLongEdge)
        {
            if (maxLongEdge <= 0)
            {
                return new Bitmap(source);
            }

            int longEdge = Math.Max(source.Width, source.Height);
            if (longEdge <= maxLongEdge)
            {
                return new Bitmap(source);
            }

            double scale = (double)maxLongEdge / longEdge;
            int width = Math.Max(1, (int)Math.Round(source.Width * scale));
            int height = Math.Max(1, (int)Math.Round(source.Height * scale));

            var resized = new Bitmap(width, height);
            using var graphics = Graphics.FromImage(resized);
            graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            graphics.DrawImage(source, new Rectangle(0, 0, width, height));
            return resized;
        }

        private static void SaveAsJpeg(Image image, Stream destination, long quality)
        {
            var encoder = ImageCodecInfo.GetImageEncoders().FirstOrDefault(codec => codec.FormatID == ImageFormat.Jpeg.Guid);
            if (encoder == null)
            {
                image.Save(destination, ImageFormat.Jpeg);
                return;
            }

            var clamped = Math.Max(1, Math.Min(100, quality));
            using var parameters = new EncoderParameters(1);
            parameters.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, clamped);
            image.Save(destination, encoder, parameters);
        }
    }
}
