using System;

namespace oPenEfficiency.Core.Services
{
    /// <summary>
    /// Host-free placement for Swap Positions (QA-S9 arrange wave): the nine anchor points
    /// of a rectangle and the exchange that moves each shape so its anchor lands where the
    /// other shape's anchor was. Production implementation used by SwapPositionsFeature.
    /// </summary>
    public static class AnchorSwapPlacementPolicy
    {
        public enum SwapAnchor
        {
            Center,
            TopLeft,
            Top,
            TopRight,
            Right,
            BottomRight,
            Bottom,
            BottomLeft,
            Left
        }

        /// <summary>The anchor point of a rectangle for the requested anchor.</summary>
        public static void GetAnchorPoint(
            SwapAnchor anchor,
            float left, float top, float width, float height,
            out float x, out float y)
        {
            switch (anchor)
            {
                case SwapAnchor.TopLeft:
                    x = left; y = top; break;
                case SwapAnchor.Top:
                    x = left + width / 2f; y = top; break;
                case SwapAnchor.TopRight:
                    x = left + width; y = top; break;
                case SwapAnchor.Right:
                    x = left + width; y = top + height / 2f; break;
                case SwapAnchor.BottomRight:
                    x = left + width; y = top + height; break;
                case SwapAnchor.Bottom:
                    x = left + width / 2f; y = top + height; break;
                case SwapAnchor.BottomLeft:
                    x = left; y = top + height; break;
                case SwapAnchor.Left:
                    x = left; y = top + height / 2f; break;
                default:
                    x = left + width / 2f; y = top + height / 2f; break;
            }
        }

        /// <summary>
        /// The swapped positions: each shape keeps its anchor-to-top-left offset but takes
        /// the other shape's anchor position, preserving size, rotation and formatting.
        /// </summary>
        public static void SwapPlacement(
            SwapAnchor anchor,
            float left1, float top1, float width1, float height1,
            float left2, float top2, float width2, float height2,
            out float newLeft1, out float newTop1, out float newLeft2, out float newTop2)
        {
            GetAnchorPoint(anchor, left1, top1, width1, height1, out float x1, out float y1);
            GetAnchorPoint(anchor, left2, top2, width2, height2, out float x2, out float y2);

            newLeft1 = x2 + (left1 - x1);
            newTop1 = y2 + (top1 - y1);
            newLeft2 = x1 + (left2 - x2);
            newTop2 = y1 + (top2 - y2);
        }
    }
}
