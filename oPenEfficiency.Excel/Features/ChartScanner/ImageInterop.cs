using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Media.Imaging;

namespace oPenEfficiency.Excel.Features.ChartScanner
{
    /// <summary>Helpers to move images between System.Drawing (capture/analysis) and WPF (display).</summary>
    internal static class ImageInterop
    {
        /// <summary>Creates a frozen WPF <see cref="BitmapSource"/> from a GDI+ bitmap (via an in-memory PNG).</summary>
        public static BitmapSource ToBitmapSource(Bitmap bitmap)
        {
            if (bitmap == null) return null;
            using (var stream = new MemoryStream())
            {
                bitmap.Save(stream, ImageFormat.Png);
                stream.Position = 0;
                var decoder = new PngBitmapDecoder(stream,
                    BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.OnLoad);
                var frame = decoder.Frames[0];
                frame.Freeze();
                return frame;
            }
        }

        /// <summary>Deep-copies a bitmap so it survives disposal of a source (e.g. clipboard image).</summary>
        public static Bitmap Clone(Image source)
        {
            if (source == null) return null;
            var bmp = new Bitmap(source.Width, source.Height, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bmp))
            {
                g.DrawImage(source, new Rectangle(0, 0, bmp.Width, bmp.Height));
            }
            return bmp;
        }
    }
}
