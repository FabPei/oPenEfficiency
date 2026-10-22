using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Shared helper for adjusting spacing between shapes (horizontal or vertical).
    /// </summary>
    internal static class AdjustSpacingFeatureHelper
    {
        public static bool Execute(
            PowerPointManager manager,
            float increment = 5f,
            bool horizontal = true,
            bool preserveCenter = false)
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

                var geometry = shapes.Select(shape => new ShapeGeometry
                {
                    Id = shape.Id,
                    Left = shape.Left,
                    Top = shape.Top,
                    Width = shape.Width,
                    Height = shape.Height
                }).ToList();
                var plannedById = ShapeGeometryPlanner
                    .AdjustSpacing(geometry, increment, horizontal, preserveCenter)
                    .ToDictionary(item => item.Id);

                foreach (var shape in shapes)
                {
                    if (horizontal)
                        shape.Left = plannedById[shape.Id].Left;
                    else
                        shape.Top = plannedById[shape.Id].Top;
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
                Utils.ExceptionLogger.Log(ex, $"AdjustSpacingFeatureHelper.Execute({(horizontal ? "H" : "V")}, Center={preserveCenter})");
                return false;
            }
        }
    }
}
