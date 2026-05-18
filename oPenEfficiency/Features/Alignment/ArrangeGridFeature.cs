using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Arrange selected shapes in a grid layout.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnArrangeGrid",
        Name = "Arrange Grid",
        Tooltip = "Arrange grid",
        IconData = "M3,3H10V10H3V3M11,3H18V10H11V3M3,11H10V18H3V11M11,11H18V18H11V11Z",
        Color = "#F43F5E",
        Description = "Arranges selected shapes in a grid layout.",
        DetailedHelpText = "### Arrange Grid\nArranges all selected shapes into an evenly-spaced grid, calculating an optimal number of rows and columns based on the selection count.",
        MinSelection = 2)]
    public static class ArrangeGridFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - arranges shapes in a 2x2 grid with 10px spacing.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, rows: 2, columns: 2, spacing: 10f);
        }

        public static bool Execute(PowerPointManager manager, int rows, int columns, float spacing = 10f)
        {
            if (manager == null) return false;
            var shapeRange = manager.GetSelectedShapes();
            if (shapeRange == null || shapeRange.Count < 2) return false;

            try
            {
                int count = shapeRange.Count;

                // Calculate bounding box of selection to find starting point
                float minLeft = float.MaxValue, minTop = float.MaxValue;
                for (int i = 1; i <= count; i++)
                {
                    if (shapeRange[i].Left < minLeft) minLeft = shapeRange[i].Left;
                    if (shapeRange[i].Top < minTop) minTop = shapeRange[i].Top;
                }

                for (int i = 0; i < count; i++)
                {
                    int r = i / columns;
                    int c = i % columns;

                    if (r >= rows) break;

                    var shape = shapeRange[i + 1];
                    shape.Left = minLeft + c * (shape.Width + spacing);
                    shape.Top = minTop + r * (shape.Height + spacing);
                }

                return true;
            }
            catch (System.Exception ex)
            {
                ExceptionLogger.Log(ex, "ArrangeGridFeature.Execute");
                return false;
            }
        }
    }
}
