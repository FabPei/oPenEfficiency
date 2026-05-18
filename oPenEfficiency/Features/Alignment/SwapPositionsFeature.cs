using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Swap positions of two shapes based on a specific anchor point.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnSwapPositions",
        Name = "Swap Positions",
        Tooltip = "Swap positions (anchor: center)",
        IconData = "M17,17L13,21V18H4V16H13V13L17,17M7,7L11,3V6H20V8H11V11L7,7Z",
        Color = "#F59E0B",
        Description = "Exchanges the X/Y coordinates of two selected shapes. Ideal for reorganization.",
        DetailedHelpText = "### Swap Positions\nSwaps the positions of two or more selected shapes. If more than two are selected, they rotate positions sequentially.\n\n**Right-Click Options:**\n* Set the **Rotation Anchor** (e.g., Top Left, Center, Bottom Right) to determine which point of the shape stays consistent during the swap.",
        MinSelection = 2,
        MaxSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class SwapPositionsFeature
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

        /// <summary>
        /// Wrapper for auto-discovery - uses Center anchor (default behavior).
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, SwapAnchor.Center);
        }

        /// <summary>
        /// Full method with parameters for manual switch usage.
        /// </summary>
        public static bool Execute(PowerPointManager manager, SwapAnchor anchor = SwapAnchor.Center)
        {
            if (manager == null) return false;
            var shapeRange = manager.GetSelectedShapes();
            if (shapeRange == null || shapeRange.Count != 2) return false;
            try
            {
                var s1 = shapeRange[1];
                var s2 = shapeRange[2];

                float x1, y1, x2, y2;
                GetAnchorPoint(s1, anchor, out x1, out y1);
                GetAnchorPoint(s2, anchor, out x2, out y2);

                // Calculate offsets from anchor to Top/Left
                float dx1 = s1.Left - x1;
                float dy1 = s1.Top - y1;
                float dx2 = s2.Left - x2;
                float dy2 = s2.Top - y2;

                // Swap
                s1.Left = x2 + dx1;
                s1.Top = y2 + dy1;
                s2.Left = x1 + dx2;
                s2.Top = y1 + dy2;

                return true;
            }
            catch (System.Exception ex)
            {
                ExceptionLogger.Log(ex, "SwapPositionsFeature.Execute");
                return false;
            }
        }

        private static void GetAnchorPoint(Shape s, SwapAnchor anchor, out float x, out float y)
        {
            switch (anchor)
            {
                case SwapAnchor.TopLeft:
                    x = s.Left; y = s.Top; break;
                case SwapAnchor.Top:
                    x = s.Left + s.Width / 2f; y = s.Top; break;
                case SwapAnchor.TopRight:
                    x = s.Left + s.Width; y = s.Top; break;
                case SwapAnchor.Right:
                    x = s.Left + s.Width; y = s.Top + s.Height / 2f; break;
                case SwapAnchor.BottomRight:
                    x = s.Left + s.Width; y = s.Top + s.Height; break;
                case SwapAnchor.Bottom:
                    x = s.Left + s.Width / 2f; y = s.Top + s.Height; break;
                case SwapAnchor.BottomLeft:
                    x = s.Left; y = s.Top + s.Height; break;
                case SwapAnchor.Left:
                    x = s.Left; y = s.Top + s.Height / 2f; break;
                case SwapAnchor.Center:
                default:
                    x = s.Left + s.Width / 2f; y = s.Top + s.Height / 2f; break;
            }
        }

        /// <summary>
        /// Internal helper to swap positions of two shapes along a single axis.
        /// </summary>
        internal static void SwapAxis(Shape s1, Shape s2, bool horizontal)
        {
            if (horizontal)
            {
                float temp = s1.Left;
                s1.Left = s2.Left;
                s2.Left = temp;
            }
            else
            {
                float temp = s1.Top;
                s1.Top = s2.Top;
                s2.Top = temp;
            }
        }
    }
}
