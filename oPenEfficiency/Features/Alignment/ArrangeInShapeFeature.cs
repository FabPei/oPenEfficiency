using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Arrange selected shapes in various shapes (Circle, Triangle, Heart, etc.).
    /// </summary>
    [FeatureMetadata(
        Id = "BtnArrangeInShape",
        Name = "Arrange in Shape",
        Tooltip = "Arrange in shape",
        IconData = "M12,2C6.48,2 2,6.48 2,12C2,17.52 6.48,22 12,22C17.52,22 22,17.52 22,12C22,6.48 17.52,2 12,2M12,4C16.42,4 20,7.58 20,12C20,16.42 16.42,20 12,20C7.58,20 4,16.42 4,12C4,7.58 7.58,4 12,4Z",
        Color = "#F43F5E",
        Description = "Arranges selected shapes in a circular, triangular, or other pattern.",
        DetailedHelpText = "### Arrange in Shape\nDistributes selected shapes into a geometric pattern within their total bounding area.\n\n**Right-Click Options:**\n* **Patterns**: Choose between Circle, Triangle, Square, Star, Octagon, or Pentagon.\n* **Pro Mode**: Opens a dialog for fine-tuned control over radius, start angle, and auto-rotation towards center.",
        MinSelection = 2)]
    public static class ArrangeInShapeFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - arranges shapes in a circle pattern.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, shapeType: "circle");
        }

        public static bool Execute(PowerPointManager manager, string shapeType)
        {
            if (manager == null) return false;
            var shapeRange = manager.GetSelectedShapes();
            if (shapeRange == null || shapeRange.Count < 2) return false;

            float minL = float.MaxValue, minT = float.MaxValue, maxR = float.MinValue, maxB = float.MinValue;
            for (int i = 1; i <= shapeRange.Count; i++)
            {
                var s = shapeRange[i];
                if (s.Left < minL) minL = s.Left;
                if (s.Top < minT) minT = s.Top;
                if (s.Left + s.Width > maxR) maxR = s.Left + s.Width;
                if (s.Top + s.Height > maxB) maxB = s.Top + s.Height;
            }
            float radius = (Math.Max(maxR - minL, maxB - minT)) / 2f;
            if (radius < 10) radius = 100;

            return ArrangeInShapePro(manager, shapeType, shapeRange.Count, radius, 0, false, false, false);
        }

        public static bool ArrangeInShapePro(PowerPointManager manager, string shapeType, int instances, float radius, float startAngleDegree, bool centerOnSlide, bool rotateToCenter, bool fillShape = false, float? anchorX = null, float? anchorY = null)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                var selection = manager.GetSelectedShapes();
                if (selection == null || selection.Count == 0) return false;

                List<Shape> shapesToArrange = new List<Shape>();
                const string InstanceTag = "oPE_ArrangePro_Instance";

                if (selection.Count == 1)
                {
                    var baseShape = selection[1];
                    shapesToArrange.Add(baseShape);

                    var slide = (Slide)baseShape.Parent;
                    var existingInstances = new List<Shape>();
                    for (int i = 1; i <= slide.Shapes.Count; i++)
                    {
                        var s = slide.Shapes[i];
                        if (s.Tags[InstanceTag] == "true") existingInstances.Add(s);
                    }

                    int needed = instances - 1;
                    while (existingInstances.Count > needed && existingInstances.Count > 0)
                    {
                        existingInstances[existingInstances.Count - 1].Delete();
                        existingInstances.RemoveAt(existingInstances.Count - 1);
                    }

                    while (existingInstances.Count < needed)
                    {
                        var dup = baseShape.Duplicate();
                        var newShape = dup[1];
                        newShape.Tags.Add(InstanceTag, "true");
                        existingInstances.Add(newShape);
                    }

                    shapesToArrange.AddRange(existingInstances);
                }
                else
                {
                    for (int i = 1; i <= selection.Count; i++) shapesToArrange.Add(selection[i]);
                }

                int count = shapesToArrange.Count;
                float centerX, centerY;
                
                if (anchorX.HasValue && anchorY.HasValue)
                {
                    centerX = anchorX.Value;
                    centerY = anchorY.Value;
                }
                else if (centerOnSlide)
                {
                    centerX = app.ActivePresentation.PageSetup.SlideWidth / 2f;
                    centerY = app.ActivePresentation.PageSetup.SlideHeight / 2f;
                }
                else
                {
                    float minL = float.MaxValue, minT = float.MaxValue, maxR = float.MinValue, maxB = float.MinValue;
                    foreach (var s in shapesToArrange)
                    {
                        if (s.Left < minL) minL = s.Left;
                        if (s.Top < minT) minT = s.Top;
                        if (s.Left + s.Width > maxR) maxR = s.Left + s.Width;
                        if (s.Top + s.Height > maxB) maxB = s.Top + s.Height;
                    }
                    centerX = (minL + maxR) / 2f;
                    centerY = (minT + maxB) / 2f;
                }

                double startAngleRad = (startAngleDegree * Math.PI) / 180.0;

                Action<Shape, float, float> applyTrans = (shape, rawX, rawY) => {
                    double dx = rawX - centerX;
                    double dy = rawY - centerY;
                    float tx = centerX + (float)(dx * Math.Cos(startAngleRad) - dy * Math.Sin(startAngleRad));
                    float ty = centerY + (float)(dx * Math.Sin(startAngleRad) + dy * Math.Cos(startAngleRad));

                    shape.Left = tx - (shape.Width / 2f);
                    shape.Top = ty - (shape.Height / 2f);

                    if (rotateToCenter) {
                        double dx2 = tx - centerX;
                        double dy2 = ty - centerY;
                        double rotAngleRad = Math.Atan2(dy2, dx2);
                        shape.Rotation = (float)(rotAngleRad * 180.0 / Math.PI) + 90; 
                    } else {
                        shape.Rotation = 0;
                    }
                };

                if (fillShape && shapeType.ToLower() == "square")
                {
                    int gridSize = (int)Math.Ceiling(Math.Sqrt(count));
                    float spacing = gridSize > 1 ? 2 * radius / (gridSize - 1) : 0;
                    int rCount = (int)Math.Ceiling((double)count / gridSize);
                    float totalHeight = (rCount - 1) * spacing;
                    float startY = centerY - totalHeight / 2f;
                    
                    int placed = 0;
                    for (int r = 0; r < rCount; r++) {
                        int itemsInRow = Math.Min(gridSize, count - placed);
                        float rowWidth = (itemsInRow - 1) * spacing;
                        float startX = centerX - rowWidth / 2f;
                        float ty_raw = startY + r * spacing;
                        
                        for (int c = 0; c < itemsInRow; c++) {
                            float tx_raw = startX + c * spacing;
                            applyTrans(shapesToArrange[placed], tx_raw, ty_raw);
                            placed++;
                        }
                    }
                }
                else if (fillShape && shapeType.ToLower() == "triangle")
                {
                    int R = (int)Math.Ceiling((-1 + Math.Sqrt(1 + 8 * count)) / 2);
                    if (R < 1) R = 1;
                    float rCircum = radius * 2.0f;
                    float triHeight = 1.5f * rCircum;
                    float maxRowWidth = (float)(rCircum * Math.Sqrt(3));
                    float spacingY = R > 1 ? triHeight / (R - 1) : 0;
                    float spacingX = R > 1 ? maxRowWidth / (R - 1) : 0;
                    float startY = centerY - rCircum; 
                    
                    int[] rowCounts = new int[R];
                    int rem = count;
                    for (int r = R - 1; r >= 0; r--) {
                        int take = Math.Min(r + 1, rem);
                        rowCounts[r] = take;
                        rem -= take;
                        if (rem <= 0) break;
                    }
                    
                    int placed = 0;
                    for (int r = 0; r < R; r++) {
                        int itemsInRow = rowCounts[r];
                        if (itemsInRow == 0) continue;
                        float rowWidth = (itemsInRow - 1) * spacingX;
                        float startX = centerX - rowWidth / 2f;
                        float ty_raw = startY + r * spacingY;
                        for (int c = 0; c < itemsInRow; c++) {
                            float tx_raw = startX + c * spacingX;
                            applyTrans(shapesToArrange[placed], tx_raw, ty_raw);
                            placed++;
                        }
                    }
                }
                else if (fillShape) 
                {
                    float goldenAngle = (float)(Math.PI * (3.0 - Math.Sqrt(5.0)));
                    for (int i = 0; i < count; i++) {
                        float r_i = radius * (count > 1 ? (float)Math.Sqrt((double)i / (count - 1)) : 0);
                        float theta = (float)startAngleRad + i * goldenAngle - (float)Math.PI/2;
                        float tx = centerX + r_i * (float)Math.Cos(theta);
                        float ty = centerY + r_i * (float)Math.Sin(theta);
                        applyTrans(shapesToArrange[i], tx, ty);
                    }
                }
                else
                {
                    // Basic shapes on outline
                    for (int i = 0; i < count; i++)
                    {
                        double angle = (2.0 * Math.PI * i / count);
                        float tx_raw = 0, ty_raw = 0;

                        switch (shapeType.ToLower())
                        {
                            case "heart":
                                float t = (float)angle;
                                tx_raw = centerX + radius * (float)(16 * Math.Pow(Math.Sin(t), 3)) / 16f;
                                ty_raw = centerY - radius * (float)(13 * Math.Cos(t) - 5 * Math.Cos(2 * t) - 2 * Math.Cos(3 * t) - Math.Cos(4 * t)) / 16f;
                                break;
                            case "triangle":
                                float t_tri = (float)angle;
                                if (t_tri < 2 * Math.PI / 3) {
                                    float p = t_tri / (float)(2 * Math.PI / 3);
                                    tx_raw = centerX + radius * (float)((1 - p) * Math.Sin(0) + p * Math.Sin(2 * Math.PI / 3));
                                    ty_raw = centerY - radius * (float)((1 - p) * Math.Cos(0) + p * Math.Cos(2 * Math.PI / 3));
                                } else if (t_tri < 4 * Math.PI / 3) {
                                    float p = (t_tri - (float)(2 * Math.PI / 3)) / (float)(2 * Math.PI / 3);
                                    tx_raw = centerX + radius * (float)((1 - p) * Math.Sin(2 * Math.PI / 3) + p * Math.Sin(4 * Math.PI / 3));
                                    ty_raw = centerY - radius * (float)((1 - p) * Math.Cos(2 * Math.PI / 3) + p * Math.Cos(4 * Math.PI / 3));
                                } else {
                                    float p = (t_tri - (float)(4 * Math.PI / 3)) / (float)(2 * Math.PI / 3);
                                    tx_raw = centerX + radius * (float)((1 - p) * Math.Sin(4 * Math.PI / 3) + p * Math.Sin(2 * Math.PI));
                                    ty_raw = centerY - radius * (float)((1 - p) * Math.Cos(4 * Math.PI / 3) + p * Math.Cos(2 * Math.PI));
                                }
                                break;
                            case "square":
                                float t_sq = (float)(angle + Math.PI / 4) % (float)(2 * Math.PI);
                                float side = (float)(2 * Math.PI / 4);
                                int sideIdx = (int)(t_sq / side);
                                float p_sq = (t_sq % side) / side;
                                float[,] corners = { {1,-1}, {1,1}, {-1,1}, {-1,-1}, {1,-1} };
                                tx_raw = centerX + radius * (corners[sideIdx, 0] * (1 - p_sq) + corners[sideIdx + 1, 0] * p_sq);
                                ty_raw = centerY + radius * (corners[sideIdx, 1] * (1 - p_sq) + corners[sideIdx + 1, 1] * p_sq);
                                break;
                            case "circle":
                            default:
                                tx_raw = centerX + radius * (float)Math.Cos(angle - Math.PI / 2);
                                ty_raw = centerY + radius * (float)Math.Sin(angle - Math.PI / 2);
                                break;
                        }
                        applyTrans(shapesToArrange[i], tx_raw, ty_raw);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ArrangeInShapeFeature.ArrangeInShapePro");
                return false;
            }
        }
    }
}
