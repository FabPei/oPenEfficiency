using System.Collections.Generic;

namespace oPenEfficiency.Excel.Features.ChartScanner
{
    /// <summary>The chart layouts the scanner can digitize.</summary>
    public enum ChartKind
    {
        /// <summary>Vertical columns: the value axis is vertical (Y).</summary>
        Columns,
        /// <summary>Horizontal bars: the value axis is horizontal (X).</summary>
        Bars,
        /// <summary>Line chart: value axis is vertical (Y); each data point is clicked individually.</summary>
        Lines
    }

    /// <summary>Orientation of the calibrated value axis.</summary>
    public enum AxisOrientation
    {
        /// <summary>Value grows along the vertical (Y) pixel axis (columns, lines).</summary>
        Vertical,
        /// <summary>Value grows along the horizontal (X) pixel axis (bars).</summary>
        Horizontal
    }

    /// <summary>A single extracted data point with its source pixel location.</summary>
    public sealed class ChartDataPoint
    {
        public string Category { get; set; }
        public double Value { get; set; }
        public double PixelX { get; set; }
        public double PixelY { get; set; }
    }

    /// <summary>The full result of a scan: chart kind and the ordered data points.</summary>
    public sealed class ChartScanResult
    {
        public ChartKind Kind { get; set; }
        public List<ChartDataPoint> Points { get; } = new List<ChartDataPoint>();
    }
}
