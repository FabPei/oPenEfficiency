using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Shared helper for adjusting spacing between shapes (horizontal or vertical).
    /// </summary>
    internal static class AdjustSpacingFeatureHelper
    {
        public static bool Execute(PowerPointManager manager, float increment = 5f, bool horizontal = true)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                var selection = app.ActiveWindow?.Selection;
                if (selection == null || selection.Type != PpSelectionType.ppSelectionShapes) return false;

                var shapeRange = manager.GetSelectedShapes();
                if (shapeRange == null || shapeRange.Count < 2) return false;

                var shapes = new List<Shape>();
                for (int i = 1; i <= shapeRange.Count; i++) shapes.Add(shapeRange[i]);

                // Sort spatially by the appropriate axis
                shapes.Sort((a, b) =>
                {
                    return horizontal ? a.Left.CompareTo(b.Left) : a.Top.CompareTo(b.Top);
                });

                // Apply cumulative offset
                for (int i = 1; i < shapes.Count; i++)
                {
                    if (horizontal)
                        shapes[i].Left += i * increment;
                    else
                        shapes[i].Top += i * increment;
                }

                // Restore selection to keep shapes selected
                try
                {
                    shapeRange.Select();
                }
                catch { /* Ignore selection restore errors */ }

                return true;
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, $"AdjustSpacingFeatureHelper.Execute({(horizontal ? "H" : "V")})");
                return false;
            }
        }
    }
}
