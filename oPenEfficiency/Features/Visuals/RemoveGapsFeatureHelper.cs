using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Shared helper for removing gaps between shapes (horizontal or vertical).
    /// </summary>
    internal static class RemoveGapsFeatureHelper
    {
        public static bool Execute(PowerPointManager manager, bool horizontal = true, bool packToCenter = false)
        {
            if (manager == null) return false;
            try
            {
                var shapeRange = manager.GetSelectedShapes();
                if (shapeRange == null || shapeRange.Count < 2) return false;

                var shapes = new List<Shape>();
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    shapes.Add(shapeRange[i]);
                }

                int? anchorId = null;
                if (!packToCenter)
                {
                    var masterShape = manager.GetReferenceShape();
                    if (masterShape == null) return false;
                    anchorId = masterShape.Id;
                }

                var geometry = shapes.Select(shape => new ShapeGeometry
                {
                    Id = shape.Id,
                    Left = shape.Left,
                    Top = shape.Top,
                    Width = shape.Width,
                    Height = shape.Height
                }).ToList();
                var plannedById = ShapeGeometryPlanner
                    .RemoveGaps(geometry, horizontal, packToCenter, anchorId)
                    .ToDictionary(item => item.Id);

                foreach (var shape in shapes)
                {
                    if (horizontal) shape.Left = plannedById[shape.Id].Left;
                    else shape.Top = plannedById[shape.Id].Top;
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, $"RemoveGapsFeatureHelper.Execute({(horizontal ? "H" : "V")})");
                return false;
            }
        }
    }
}
