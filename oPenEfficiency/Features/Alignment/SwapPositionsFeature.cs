using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using SwapAnchor = oPenEfficiency.Core.Services.AnchorSwapPlacementPolicy.SwapAnchor;

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
        DetailedHelpText = "### Swap Positions\n\nSelect exactly two shapes to exchange their positions. The standard action swaps their centers. Right-click to instead exchange matching top, side, corner, or bottom anchor points. Shape size, rotation, and formatting remain unchanged.",
        MinSelection = 2,
        MaxSelection = 2,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class SwapPositionsFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - uses Center anchor (default behavior).
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, SwapAnchor.Center);
        }

        /// <summary>
        /// Full method with parameters for explicit command variants.
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
                float originalLeft1 = s1.Left, originalTop1 = s1.Top;
                float originalLeft2 = s2.Left, originalTop2 = s2.Top;

                oPenEfficiency.Core.Services.AnchorSwapPlacementPolicy.SwapPlacement(
                    anchor,
                    s1.Left, s1.Top, s1.Width, s1.Height,
                    s2.Left, s2.Top, s2.Width, s2.Height,
                    out float newLeft1, out float newTop1, out float newLeft2, out float newTop2);

                try { manager.GetApplication()?.StartNewUndoEntry(); } catch { /* Swap remains valid when undo grouping is unavailable. */ }
                try
                {
                    s1.Left = newLeft1;
                    s1.Top = newTop1;
                    s2.Left = newLeft2;
                    s2.Top = newTop2;
                }
                catch
                {
                    try { s1.Left = originalLeft1; s1.Top = originalTop1; } catch { /* Preserve the primary swap error if the first proxy is stale. */ }
                    try { s2.Left = originalLeft2; s2.Top = originalTop2; } catch { /* Preserve the primary swap error if the second proxy is stale. */ }
                    throw;
                }

                return true;
            }
            catch (System.Exception ex)
            {
                ExceptionLogger.Log(ex, "SwapPositionsFeature.Execute");
                return false;
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
