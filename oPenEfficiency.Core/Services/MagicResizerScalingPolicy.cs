using System;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free scaling contract for Scale Selected Shapes (QA-S9 align-and-size wave):
    /// the centimetre/point conversions, the percentage and absolute scale-factor
    /// derivations, the nine anchor fractions, the anchor-preserving per-shape bounds and
    /// the scalar factors for line weight, bullet size, margins and font size.
    /// Production implementation used by MagicResizerWindow.
    /// </summary>
    public static class MagicResizerScalingPolicy
    {
        /// <summary>The nine anchor points of the 3x3 anchor grid (fractions of the shape bounds).</summary>
        public enum ScaleAnchor
        {
            TopLeft,
            Top,
            TopRight,
            Left,
            Center,
            Right,
            BottomLeft,
            Bottom,
            BottomRight
        }

        /// <summary>Points to centimetres (PowerPoint stores points; the dialog asks for centimetres).</summary>
        public static double PointsToCentimetres(float points)
        {
            return points * 2.54 / 72.0;
        }

        /// <summary>Centimetres to points.</summary>
        public static double CentimetresToPoints(double centimetres)
        {
            return centimetres * 72.0 / 2.54;
        }

        /// <summary>
        /// Percentage mode: both factors are the percentage divided by 100. A non-positive
        /// percentage is rejected (the caller shows the historical warning).
        /// </summary>
        public static bool TryPercentageScaleFactors(double percentage, out double scaleX, out double scaleY)
        {
            scaleX = 1.0;
            scaleY = 1.0;
            if (percentage <= 0) return false;

            scaleX = percentage / 100.0;
            scaleY = percentage / 100.0;
            return true;
        }

        /// <summary>
        /// Absolute mode: the target width/height in centimetres are converted to points and
        /// divided by the reference shape's own dimensions, so every selected shape scales by
        /// the same factors the reference shape would need to reach the requested size.
        /// Non-positive targets or reference dimensions are rejected.
        /// </summary>
        public static bool TryAbsoluteScaleFactors(
            double widthCm, double heightCm,
            float masterWidthPts, float masterHeightPts,
            out double scaleX, out double scaleY)
        {
            scaleX = 1.0;
            scaleY = 1.0;
            if (widthCm <= 0 || heightCm <= 0 || masterWidthPts <= 0 || masterHeightPts <= 0) return false;

            double wTargetPts = CentimetresToPoints(widthCm);
            double hTargetPts = CentimetresToPoints(heightCm);
            scaleX = wTargetPts / masterWidthPts;
            scaleY = hTargetPts / masterHeightPts;
            return true;
        }

        /// <summary>The anchor fractions (0 = left/top, 0.5 = center, 1 = right/bottom) of the anchor point.</summary>
        public static void GetAnchorFractions(ScaleAnchor anchor, out double anchorX, out double anchorY)
        {
            switch (anchor)
            {
                case ScaleAnchor.TopLeft:
                    anchorX = 0.0; anchorY = 0.0; break;
                case ScaleAnchor.Top:
                    anchorX = 0.5; anchorY = 0.0; break;
                case ScaleAnchor.TopRight:
                    anchorX = 1.0; anchorY = 0.0; break;
                case ScaleAnchor.Left:
                    anchorX = 0.0; anchorY = 0.5; break;
                case ScaleAnchor.Right:
                    anchorX = 1.0; anchorY = 0.5; break;
                case ScaleAnchor.BottomLeft:
                    anchorX = 0.0; anchorY = 1.0; break;
                case ScaleAnchor.Bottom:
                    anchorX = 0.5; anchorY = 1.0; break;
                case ScaleAnchor.BottomRight:
                    anchorX = 1.0; anchorY = 1.0; break;
                default:
                    anchorX = 0.5; anchorY = 0.5; break;
            }
        }

        /// <summary>True when both factors are exactly 1.0 — the dialog then changes nothing.</summary>
        public static bool IsNoOp(double scaleX, double scaleY)
        {
            return scaleX == 1.0 && scaleY == 1.0;
        }

        /// <summary>
        /// The scaled bounds of one shape: the anchor point stays where it was, so the new
        /// top-left is the fixed point minus the scaled offset of the anchor within the shape.
        /// </summary>
        public static void ComputeScaledBounds(
            float left, float top, float width, float height,
            double scaleX, double scaleY, ScaleAnchor anchor,
            out float newLeft, out float newTop, out float newWidth, out float newHeight)
        {
            GetAnchorFractions(anchor, out double anchorX, out double anchorY);

            // Fixed point based on anchor
            float fixedX = left + (width * (float)anchorX);
            float fixedY = top + (height * (float)anchorY);

            // Resize
            newWidth = width * (float)scaleX;
            newHeight = height * (float)scaleY;

            // Reposition to maintain anchor
            // fixedX = newLeft + (newWidth * anchorX)  =>  newLeft = fixedX - (newWidth * anchorX)
            newLeft = fixedX - (newWidth * (float)anchorX);
            newTop = fixedY - (newHeight * (float)anchorY);
        }

        /// <summary>The scalar factor used for the properties that follow one average of both scale factors.</summary>
        public static double ScalarFactor(double scaleX, double scaleY)
        {
            return (scaleX + scaleY) / 2.0;
        }

        /// <summary>Scales a scalar property (line weight, bullet relative size, text-frame margins) by one factor.</summary>
        public static float ScaleScalar(float value, double factor)
        {
            return value * (float)factor;
        }

        /// <summary>The scaled font size, floored at 1 pt (the historical minimum).</summary>
        public static float ScaledFontSize(float fontSize, double factor)
        {
            float newSize = fontSize * (float)factor;
            if (newSize < 1) newSize = 1;
            return newSize;
        }
    }
}
