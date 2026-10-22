using System;

namespace oPenEfficiency.Excel.Features.ChartScanner
{
    /// <summary>
    /// Linear pixel-to-value mapping along one axis, defined by two reference points
    /// with known values (the "relative scaling" the user calibrates by hand).
    ///
    /// This is pure mathematics with no WPF/Excel dependency, so it is fully
    /// deterministic and unit-testable. Pixel coordinates use the usual screen
    /// convention where Y grows downward; the linear fit handles that transparently.
    /// </summary>
    public sealed class AxisCalibration
    {
        public AxisOrientation Orientation { get; }
        public double Pixel1 { get; }
        public double Value1 { get; }
        public double Pixel2 { get; }
        public double Value2 { get; }

        public AxisCalibration(AxisOrientation orientation,
                               double pixel1, double value1,
                               double pixel2, double value2)
        {
            Orientation = orientation;
            Pixel1 = pixel1;
            Value1 = value1;
            Pixel2 = pixel2;
            Value2 = value2;
        }

        /// <summary>True if the two reference points are distinct enough to define a scale.</summary>
        public bool IsValid => Math.Abs(Pixel2 - Pixel1) > 1e-6;

        /// <summary>
        /// Maps a pixel position along the calibrated axis to a data value.
        /// For vertical axes pass the pixel Y; for horizontal axes pass the pixel X.
        /// </summary>
        public double ToValue(double pixel)
        {
            if (!IsValid) return double.NaN;
            double t = (pixel - Pixel1) / (Pixel2 - Pixel1);
            return Value1 + t * (Value2 - Value1);
        }

        /// <summary>Picks the pixel coordinate that this axis maps from, for a point.</summary>
        public double PixelOf(double pixelX, double pixelY)
            => Orientation == AxisOrientation.Vertical ? pixelY : pixelX;

        /// <summary>Convenience: maps a 2D pixel location to a value using this axis' orientation.</summary>
        public double ValueAt(double pixelX, double pixelY)
            => ToValue(PixelOf(pixelX, pixelY));
    }
}
