using System;
using System.Collections.Generic;
using System.Drawing;

namespace oPenEfficiency.Excel.Features.ChartScanner
{
    /// <summary>
    /// Colour-based tracing of a single line series. The user samples the line colour; the
    /// tracer builds the same match mask as <see cref="BarDetector"/>, finds the line's
    /// horizontal extent, then samples the vertical centroid at <c>count</c> evenly spaced
    /// X positions – one proposed data point each. Results are corrigible in the digitizer.
    ///
    /// Points are returned as (X, centroid-Y) so <see cref="AxisCalibration.ValueAt"/> reads
    /// the value off the vertical value axis (line charts use a vertical value axis).
    /// </summary>
    internal static class LineTracer
    {
        public static List<Point> Trace(Bitmap image, Color color, double tolerance, int count)
        {
            var points = new List<Point>();
            if (image == null || count < 1) return points;

            int w = image.Width, h = image.Height;
            bool[,] mask = BarDetector.BuildMatchMask(image, color, tolerance);

            // Horizontal extent of the traced line.
            int minX = -1, maxX = -1;
            for (int x = 0; x < w; x++)
            {
                if (!ColumnHasMatch(mask, h, x)) continue;
                if (minX < 0) minX = x;
                maxX = x;
            }
            if (minX < 0) return points;

            // Half-window used to average out anti-aliasing / small gaps at each sample X.
            int span = Math.Max(1, maxX - minX);
            int win = Math.Max(1, span / (count * 4));

            for (int i = 0; i < count; i++)
            {
                double tx = count == 1 ? (minX + maxX) / 2.0 : minX + (double)span * i / (count - 1);
                int cx = (int)Math.Round(tx);

                double sumY = 0; int n = 0;
                for (int x = Math.Max(0, cx - win); x <= Math.Min(w - 1, cx + win); x++)
                    for (int y = 0; y < h; y++)
                        if (mask[x, y]) { sumY += y; n++; }

                if (n == 0) continue;
                int cy = (int)Math.Round(sumY / n);
                points.Add(new Point(cx, cy));
            }

            return points;
        }

        private static bool ColumnHasMatch(bool[,] mask, int h, int x)
        {
            for (int y = 0; y < h; y++) if (mask[x, y]) return true;
            return false;
        }
    }
}
