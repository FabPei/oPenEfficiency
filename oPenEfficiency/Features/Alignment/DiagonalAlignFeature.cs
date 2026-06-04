using System;
using System.Linq;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Alignment
{
    [FeatureMetadata(Id = "BtnDiagonalAlign", Name = "Diagonal Align", Tooltip = "Align selection diagonally at a specific degree", IconData = "M21 3L3 21M21 3l-6 0M21 3l0 6", Color = "#8B5CF6", Description = "Aligns selected shapes in a diagonal line (e.g., 45 degrees). Use the right-click menu to choose specific angles.", Keywords = "diagonal alignment, staircase layout, angular distribute", MinSelection = 2, RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class DiagonalAlignFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try
            {
                var window = new UI.Toolbars.DiagonalAlignToolbar(manager);
                window.Show();
                return true;
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "DiagonalAlignFeature.Execute");
                return false;
            }
        }

        public static bool ExecuteSpecific(PowerPointManager manager, float angleDeg, string anchorType = "First")
        {
            var app = manager?.GetApplication();
            if (app?.ActiveWindow?.Selection?.Type != PpSelectionType.ppSelectionShapes) return false;
            
            var sr = app.ActiveWindow.Selection.ShapeRange;
            if (sr.Count < 2) return false;

            try
            {
                var shapes = new List<ShapeData>();
                for (int i = 1; i <= sr.Count; i++)
                {
                    var s = sr[i];
                    shapes.Add(new ShapeData
                    {
                        Shape = s,
                        OrigCenterX = s.Left + (s.Width / 2f),
                        OrigCenterY = s.Top + (s.Height / 2f),
                        OrigWidth = s.Width,
                        OrigHeight = s.Height
                    });
                }

                var sortedShapes = shapes.OrderBy(s => s.OrigCenterX + s.OrigCenterY).ToList();
                var first = sortedShapes.First();
                var last = sortedShapes.Last();

                float dx = last.OrigCenterX - first.OrigCenterX;
                float dy = last.OrigCenterY - first.OrigCenterY;
                float dist = (float)Math.Sqrt(dx * dx + dy * dy);
                float spacing = dist / (shapes.Count - 1);

                float angleRad = (float)(angleDeg * Math.PI / 180f);
                float dirX = (float)Math.Cos(angleRad);
                float dirY = (float)Math.Sin(angleRad);

                float avgX = shapes.Average(s => s.OrigCenterX);
                float avgY = shapes.Average(s => s.OrigCenterY);

                var projectedShapes = shapes.Select(s => new
                {
                    Data = s,
                    Proj = (s.OrigCenterX - avgX) * dirX + (s.OrigCenterY - avgY) * dirY
                }).OrderBy(x => x.Proj).ToList();

                float anchorOrigX = avgX;
                float anchorOrigY = avgY;
                float anchorIndex = (projectedShapes.Count - 1) / 2f;

                if (anchorType == "First")
                {
                    anchorOrigX = projectedShapes.First().Data.OrigCenterX;
                    anchorOrigY = projectedShapes.First().Data.OrigCenterY;
                    anchorIndex = 0;
                }
                else if (anchorType == "Last")
                {
                    anchorOrigX = projectedShapes.Last().Data.OrigCenterX;
                    anchorOrigY = projectedShapes.Last().Data.OrigCenterY;
                    anchorIndex = projectedShapes.Count - 1;
                }

                for (int i = 0; i < projectedShapes.Count; i++)
                {
                    var item = projectedShapes[i];
                    float offset = (i - anchorIndex) * spacing;
                    
                    float newCenterX = anchorOrigX + offset * dirX;
                    float newCenterY = anchorOrigY + offset * dirY;

                    try
                    {
                        item.Data.Shape.Left = newCenterX - (item.Data.OrigWidth / 2f);
                        item.Data.Shape.Top = newCenterY - (item.Data.OrigHeight / 2f);
                    }
                    catch { } // Ignore if shape was deleted
                }

                return true;
            }
            catch (Exception ex)
            {
                Utils.ExceptionLogger.Log(ex, "DiagonalAlignFeature.ExecuteSpecific");
                return false;
            }
        }

        private class ShapeData
        {
            public Microsoft.Office.Interop.PowerPoint.Shape Shape { get; set; }
            public float OrigCenterX { get; set; }
            public float OrigCenterY { get; set; }
            public float OrigWidth { get; set; }
            public float OrigHeight { get; set; }
        }
    }
}