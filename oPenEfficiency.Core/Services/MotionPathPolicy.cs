using System;
using System.Globalization;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free motion-path contract for Create Motion Path (QA-S9 arrange wave): which
    /// directions a mode expands to, the slide-relative delta between the two shapes'
    /// centers and the invariant-culture motion path string PowerPoint receives.
    /// Production implementation used by CreateMotionPathFeature.
    /// </summary>
    public static class MotionPathPolicy
    {
        /// <summary>Which direction(s) the command animates.</summary>
        public enum MotionPathMode
        {
            FirstToSecond,
            SecondToFirst,
            BothDirections
        }

        /// <summary>True when the mode animates from the first selected shape to the second.</summary>
        public static bool MovesFirstToSecond(MotionPathMode mode)
        {
            return mode == MotionPathMode.FirstToSecond || mode == MotionPathMode.BothDirections;
        }

        /// <summary>True when the mode animates from the second selected shape back to the first.</summary>
        public static bool MovesSecondToFirst(MotionPathMode mode)
        {
            return mode == MotionPathMode.SecondToFirst || mode == MotionPathMode.BothDirections;
        }

        /// <summary>
        /// The center-to-center delta of the two shapes as a fraction of the slide size
        /// (the unit PowerPoint's motion path behavior expects).
        /// </summary>
        public static void ComputeNormalizedDelta(
            float fromLeft, float fromTop, float fromWidth, float fromHeight,
            float toLeft, float toTop, float toWidth, float toHeight,
            float slideWidth, float slideHeight,
            out float pctX, out float pctY)
        {
            float fromCenterX = fromLeft + fromWidth / 2f;
            float fromCenterY = fromTop + fromHeight / 2f;
            float toCenterX = toLeft + toWidth / 2f;
            float toCenterY = toTop + toHeight / 2f;

            pctX = (toCenterX - fromCenterX) / slideWidth;
            pctY = (toCenterY - fromCenterY) / slideHeight;
        }

        /// <summary>
        /// The motion path string, always with invariant-culture coordinates regardless of
        /// the user's locale.
        /// </summary>
        public static string BuildPathString(float pctX, float pctY)
        {
            return string.Format(
                CultureInfo.InvariantCulture,
                "M 0 0 L {0} {1} E",
                pctX.ToString(CultureInfo.InvariantCulture),
                pctY.ToString(CultureInfo.InvariantCulture));
        }
    }
}
