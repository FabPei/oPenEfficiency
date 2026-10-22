using System;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free placement for Snap Shape to Object (QA-S9 arrange wave): the nearest-side
    /// direction decision from the two shape centers and the flush placement centered on the
    /// perpendicular axis. Production implementation used by SnapShapeFeature.
    /// </summary>
    public static class ShapeSnapPlacementPolicy
    {
        public enum SnapDirection
        {
            Nearest,
            Left,
            Right,
            Above,
            Below
        }

        /// <summary>
        /// The nearest side of the target rectangle as seen from the moving shape's center:
        /// the horizontal axis wins ties, decided by the sign of the center delta.
        /// </summary>
        public static SnapDirection GetNearestDirection(
            float movingLeft, float movingTop, float movingWidth, float movingHeight,
            float targetLeft, float targetTop, float targetWidth, float targetHeight)
        {
            float deltaX = movingLeft + movingWidth / 2f - (targetLeft + targetWidth / 2f);
            float deltaY = movingTop + movingHeight / 2f - (targetTop + targetHeight / 2f);
            if (Math.Abs(deltaX) >= Math.Abs(deltaY)) return deltaX < 0 ? SnapDirection.Left : SnapDirection.Right;
            return deltaY < 0 ? SnapDirection.Above : SnapDirection.Below;
        }

        /// <summary>
        /// Places the moving shape flush to the target on the requested side, centered on the
        /// perpendicular axis. An unknown direction anchors at the target's top-left.
        /// </summary>
        public static void Place(
            SnapDirection direction,
            float targetLeft, float targetTop, float targetWidth, float targetHeight,
            float movingWidth, float movingHeight,
            out float left, out float top)
        {
            switch (direction)
            {
                case SnapDirection.Left:
                    left = targetLeft - movingWidth;
                    top = targetTop + (targetHeight - movingHeight) / 2f;
                    break;
                case SnapDirection.Right:
                    left = targetLeft + targetWidth;
                    top = targetTop + (targetHeight - movingHeight) / 2f;
                    break;
                case SnapDirection.Above:
                    left = targetLeft + (targetWidth - movingWidth) / 2f;
                    top = targetTop - movingHeight;
                    break;
                case SnapDirection.Below:
                    left = targetLeft + (targetWidth - movingWidth) / 2f;
                    top = targetTop + targetHeight;
                    break;
                default:
                    left = targetLeft;
                    top = targetTop;
                    break;
            }
        }
    }
}
