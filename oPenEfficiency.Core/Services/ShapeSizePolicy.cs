using System;
using System.Globalization;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free decision rules behind the Shape Size Panel widget (QA-S9 wave 3): the centimetre
    /// point conversion, the typed dimension parsing (with the pt/cm suffixes and the culture
    /// fallback chain), the typed-input validity, the step-adjust floor, the per-shape change
    /// detection and the uniform-size display rule. Production implementation used by
    /// ShapeSizePanel.
    /// </summary>
    public static class ShapeSizePolicy
    {
        public const double PointsPerCentimetre = 28.3464567;

        /// <summary>Smallest adjustable dimension in points (0.01 cm).</summary>
        public static double MinimumAdjustPoints => CmToPoints(0.01);

        /// <summary>Skip tolerance for applying a new dimension (points).</summary>
        public const float ApplyEpsilonPoints = 0.001f;

        /// <summary>Tolerance for treating several shapes as uniformly sized (points).</summary>
        public const float UniformEpsilonPoints = 0.01f;

        public static double CmToPoints(double cm) => cm * PointsPerCentimetre;

        public static double PointsToCm(double points) => points / PointsPerCentimetre;

        /// <summary>Formats a centimetre value exactly as the panel displays it.</summary>
        public static string FormatCm(double cm)
        {
            return cm.ToString("0.##", CultureInfo.CurrentCulture) + " cm";
        }

        /// <summary>
        /// Parses typed dimension text: optional "pt"/"cm" suffix (points are converted to cm),
        /// then the user's culture, the invariant culture, and finally the invariant culture with
        /// a comma replaced by a point. NaN/infinity are rejected.
        /// </summary>
        public static bool TryParseDisplay(string text, out double cm)
        {
            cm = 0;
            string normalized = (text ?? string.Empty).Trim();
            bool isPoints = normalized.EndsWith("pt", StringComparison.OrdinalIgnoreCase);
            if (isPoints) normalized = normalized.Substring(0, normalized.Length - 2).Trim();
            else if (normalized.EndsWith("cm", StringComparison.OrdinalIgnoreCase))
                normalized = normalized.Substring(0, normalized.Length - 2).Trim();

            if (!double.TryParse(normalized, NumberStyles.Float, CultureInfo.CurrentCulture, out double value) &&
                !double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out value) &&
                !double.TryParse(normalized.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out value))
                return false;

            cm = isPoints ? PointsToCm(value) : value;
            return !double.IsNaN(cm) && !double.IsInfinity(cm);
        }

        /// <summary>Typed dimensions must be positive (NaN fails this as well).</summary>
        public static bool IsValidTypedCm(double cm)
        {
            return cm > 0;
        }

        /// <summary>The stepped target: current size plus the delta, floored at the 0.01 cm minimum.</summary>
        public static double TargetAfterAdjust(double currentPoints, double deltaCm)
        {
            return Math.Max(MinimumAdjustPoints, currentPoints + CmToPoints(deltaCm));
        }

        /// <summary>Whether a shape should receive a new dimension (valid target, meaningfully different).</summary>
        public static bool NeedsDimensionChange(float currentPoints, float targetPoints)
        {
            if (targetPoints <= 0f || float.IsNaN(targetPoints) || float.IsInfinity(targetPoints)) return false;
            return Math.Abs(targetPoints - currentPoints) >= ApplyEpsilonPoints;
        }

        /// <summary>Whether two shapes count as uniformly sized for the panel display.</summary>
        public static bool AreUniform(float first, float other)
        {
            return Math.Abs(other - first) <= UniformEpsilonPoints;
        }
    }
}
