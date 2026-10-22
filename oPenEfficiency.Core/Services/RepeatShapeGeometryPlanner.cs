using System;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free geometry for Repeat Shape (QA-S9 wave 5): the gap inference between the two
    /// selected rectangles and the placement of the duplicate in the repeat direction. Production
    /// implementation used by RepeatShapeFeature.
    /// </summary>
    public static class RepeatShapeGeometryPlanner
    {
        public enum RepeatPlacementDirection
        {
            Right,
            Left,
            Up,
            Down
        }

        /// <summary>
        /// Infers the horizontal gap the user's two selected shapes imply: the edge-to-edge gap
        /// when the rectangles are horizontally disjoint, otherwise the absolute left offset.
        /// </summary>
        public static float InferGapX(float left1, float width1, float left2, float width2)
        {
            float right1 = left1 + width1;
            float right2 = left2 + width2;
            if (Math.Max(left1, left2) > Math.Min(right1, right2))
                return Math.Max(left1, left2) - Math.Min(right1, right2);
            return Math.Abs(left1 - left2);
        }

        /// <summary>
        /// Infers the vertical gap the user's two selected shapes imply: the edge-to-edge gap
        /// when the rectangles are vertically disjoint, otherwise the absolute top offset.
        /// </summary>
        public static float InferGapY(float top1, float height1, float top2, float height2)
        {
            float bottom1 = top1 + height1;
            float bottom2 = top2 + height2;
            if (Math.Max(top1, top2) > Math.Min(bottom1, bottom2))
                return Math.Max(top1, top2) - Math.Min(bottom1, bottom2);
            return Math.Abs(top1 - top2);
        }

        /// <summary>
        /// Places the duplicate next to the anchor shape (the second selected shape) in the given
        /// direction. Left/Up placement uses the duplicate's own width/height, matching the
        /// historical behavior.
        /// </summary>
        public static void NextPlacement(
            RepeatPlacementDirection direction,
            float anchorLeft,
            float anchorTop,
            float anchorWidth,
            float anchorHeight,
            float newWidth,
            float newHeight,
            float gapX,
            float gapY,
            out float newLeft,
            out float newTop)
        {
            switch (direction)
            {
                case RepeatPlacementDirection.Right:
                    newLeft = anchorLeft + anchorWidth + gapX;
                    newTop = anchorTop;
                    break;
                case RepeatPlacementDirection.Left:
                    newLeft = anchorLeft - newWidth - gapX;
                    newTop = anchorTop;
                    break;
                case RepeatPlacementDirection.Down:
                    newTop = anchorTop + anchorHeight + gapY;
                    newLeft = anchorLeft;
                    break;
                default:
                    newTop = anchorTop - newHeight - gapY;
                    newLeft = anchorLeft;
                    break;
            }
        }
    }
}
