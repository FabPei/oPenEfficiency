using System;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free pixel rule behind the Transparent Color feature (QA-S9 wave 3): given the sampled
    /// color and a tolerance percentage, an opaque pixel is made transparent when its squared
    /// Euclidean RGB distance to the sample stays within the squared Cartesian limit
    /// 3 * 255^2 * tolerance^2. Production implementation used by TransparentColorWindow.
    /// </summary>
    /// <remarks>
    /// The boundary is inclusive (distSq == limitSq becomes transparent) and already transparent
    /// pixels (alpha 0) are never touched - both behaviors are part of the pinned contract.
    /// </remarks>
    public static class PixelTransparencyPolicy
    {
        /// <summary>
        /// The squared distance limit for a tolerance percentage in 0..1. Mirrors the production
        /// arithmetic exactly (double math narrowed to float): the rounding at the boundary is part
        /// of the pinned contract, so the limit must stay single precision.
        /// </summary>
        public static float LimitSquared(double tolerancePercent)
        {
            double tolerance = Math.Max(0d, Math.Min(1d, tolerancePercent));
            return (float)(3d * 255d * 255d * (tolerance * tolerance));
        }

        /// <summary>
        /// Whether the pixel should become transparent: it must be opaque (alpha > 0) and its
        /// squared RGB distance to the sampled color must stay at or below the tolerance limit.
        /// </summary>
        public static bool ShouldMakeTransparent(
            int red,
            int green,
            int blue,
            int alpha,
            int targetRed,
            int targetGreen,
            int targetBlue,
            double tolerancePercent)
        {
            if (alpha <= 0) return false;

            int redDelta = red - targetRed;
            int greenDelta = green - targetGreen;
            int blueDelta = blue - targetBlue;
            int distanceSquared = redDelta * redDelta +
                                  greenDelta * greenDelta +
                                  blueDelta * blueDelta;

            return distanceSquared <= LimitSquared(tolerancePercent);
        }
    }
}
