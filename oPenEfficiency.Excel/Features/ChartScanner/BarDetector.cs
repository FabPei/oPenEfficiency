using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;

namespace oPenEfficiency.Excel.Features.ChartScanner
{
    /// <summary>
    /// Colour-based automatic detection of bars/columns. The user samples one bar's colour;
    /// the detector then finds every contiguous run of matching pixels and returns one tip
    /// point per bar. The results are proposals the user can correct in the digitizer.
    ///
    /// Returned points are already oriented for <see cref="AxisCalibration.ValueAt"/>:
    ///   Columns -> (centre X, bar-top Y)      Bars -> (bar-end X, centre Y)
    /// so the caller can compute the value directly from the calibrated value axis.
    /// </summary>
    internal static class BarDetector
    {
        /// <summary>Detects bars of the given colour. <paramref name="tolerance"/> is a per-pixel RGB distance.</summary>
        public static List<Point> Detect(Bitmap image, Color target, double tolerance, ChartKind kind)
        {
            var points = new List<Point>();
            if (image == null) return points;

            int w = image.Width, h = image.Height;
            bool[,] match = BuildMatchMask(image, target, tolerance);

            if (kind == ChartKind.Bars)
            {
                // Horizontal bars: group consecutive matching ROWS; tip is the rightmost pixel.
                int minFill = Math.Max(3, w / 200);
                int y = 0;
                while (y < h)
                {
                    if (RowFill(match, w, y) < minFill) { y++; continue; }
                    int start = y;
                    while (y < h && RowFill(match, w, y) >= minFill) y++;
                    int end = y - 1;
                    int centerY = (start + end) / 2;
                    int tipX = 0;
                    for (int yy = start; yy <= end; yy++)
                        for (int xx = w - 1; xx >= 0; xx--)
                            if (match[xx, yy]) { if (xx > tipX) tipX = xx; break; }
                    points.Add(new Point(tipX, centerY));
                }
            }
            else
            {
                // Columns (and default): group consecutive matching COLUMNS; tip is the topmost pixel.
                int minFill = Math.Max(3, h / 200);
                int x = 0;
                while (x < w)
                {
                    if (ColFill(match, h, x) < minFill) { x++; continue; }
                    int start = x;
                    while (x < w && ColFill(match, h, x) >= minFill) x++;
                    int end = x - 1;
                    int centerX = (start + end) / 2;
                    int tipY = h;
                    for (int xx = start; xx <= end; xx++)
                        for (int yy = 0; yy < h; yy++)
                            if (match[xx, yy]) { if (yy < tipY) tipY = yy; break; }
                    points.Add(new Point(centerX, tipY));
                }
            }

            return points;
        }

        /// <summary>Builds a per-pixel colour-match mask (shared with the line tracer).</summary>
        internal static bool[,] BuildMatchMask(Bitmap image, Color target, double tolerance)
        {
            int w = image.Width, h = image.Height;
            var mask = new bool[w, h];
            double tol2 = tolerance * tolerance;

            var rect = new Rectangle(0, 0, w, h);
            BitmapData data = image.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            try
            {
                int stride = data.Stride;
                int bytes = Math.Abs(stride) * h;
                var buffer = new byte[bytes];
                System.Runtime.InteropServices.Marshal.Copy(data.Scan0, buffer, 0, bytes);

                for (int y = 0; y < h; y++)
                {
                    int row = y * stride;
                    for (int x = 0; x < w; x++)
                    {
                        int i = row + x * 4;      // BGRA
                        double db = buffer[i] - target.B;
                        double dg = buffer[i + 1] - target.G;
                        double dr = buffer[i + 2] - target.R;
                        double a = buffer[i + 3];
                        if (a > 32 && (dr * dr + dg * dg + db * db) <= tol2)
                            mask[x, y] = true;
                    }
                }
            }
            finally { image.UnlockBits(data); }
            return mask;
        }

        private static int ColFill(bool[,] mask, int h, int x)
        {
            int n = 0;
            for (int y = 0; y < h; y++) if (mask[x, y]) n++;
            return n;
        }

        private static int RowFill(bool[,] mask, int w, int y)
        {
            int n = 0;
            for (int x = 0; x < w; x++) if (mask[x, y]) n++;
            return n;
        }
    }
}
