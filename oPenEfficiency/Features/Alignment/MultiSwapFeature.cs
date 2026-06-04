using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnMultiSwap",
        Name = "Multi-Swap",
        Tooltip = "Swap multiple shapes based on selection order",
        IconData = "M16,3H21V8 M4,20L21,3 M21,16V21H16 M15,15L21,21 M4,4L9,9",
        Color = "#F59E0B",
        Description = "Swaps positions of multiple shapes based on their selection order and a chosen direction.",
        DetailedHelpText = "### Multi-Swap\nSwaps the positions of multiple selected shapes. The first selected shape moves to the primary position (e.g., top-most), the second to the next, and so on.\n\n**Right-Click Options:**\n* **Top-Down:** Sorts target positions from top to bottom.\n* **Bottom-Up:** Sorts target positions from bottom to top.\n* **Left-Right:** Sorts target positions from left to right.\n* **Right-Left:** Sorts target positions from right to left.",
        Keywords = "swap positions, rearrange, flip locations, switch places",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class MultiSwapFeature
    {
        public enum SwapDirection
        {
            TopDown,
            BottomUp,
            LeftRight,
            RightLeft
        }

        /// <summary>
        /// Wrapper for auto-discovery - defaults to Top-Down direction.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, SwapDirection.TopDown);
        }

        /// <summary>
        /// Executes the multi-swap based on the specified direction.
        /// </summary>
        public static bool Execute(PowerPointManager manager, SwapDirection direction)
        {
            if (manager == null) return false;
            var shapeRange = manager.GetSelectedShapes();
            if (shapeRange == null || shapeRange.Count < 2) return false;

            try
            {
                // 1. Capture shapes in selection order
                // Note: PowerPoint's ShapeRange index is 1-based.
                // We rely on the fact that ShapeRange typically reflects selection order.
                var selectedShapes = new List<Shape>();
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    selectedShapes.Add(shapeRange[i]);
                }

                // 2. Capture original positions
                var originalPositions = selectedShapes.Select(s => new { s.Left, s.Top }).ToList();

                // 3. Sort target positions based on direction
                List<dynamic> sortedPositions;
                switch (direction)
                {
                    case SwapDirection.BottomUp:
                        sortedPositions = originalPositions.OrderByDescending(p => p.Top).ThenBy(p => p.Left).Cast<dynamic>().ToList();
                        break;
                    case SwapDirection.LeftRight:
                        sortedPositions = originalPositions.OrderBy(p => p.Left).ThenBy(p => p.Top).Cast<dynamic>().ToList();
                        break;
                    case SwapDirection.RightLeft:
                        sortedPositions = originalPositions.OrderByDescending(p => p.Left).ThenBy(p => p.Top).Cast<dynamic>().ToList();
                        break;
                    case SwapDirection.TopDown:
                    default:
                        sortedPositions = originalPositions.OrderBy(p => p.Top).ThenBy(p => p.Left).Cast<dynamic>().ToList();
                        break;
                }

                // 4. Apply positions back to shapes in selection order
                for (int i = 0; i < selectedShapes.Count; i++)
                {
                    selectedShapes[i].Left = sortedPositions[i].Left;
                    selectedShapes[i].Top = sortedPositions[i].Top;
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "MultiSwapFeature.Execute");
                return false;
            }
        }
    }
}
