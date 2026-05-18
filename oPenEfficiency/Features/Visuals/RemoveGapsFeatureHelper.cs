using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
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

                if (horizontal)
                {
                    shapes.Sort((a, b) => a.Left.CompareTo(b.Left));

                    if (packToCenter)
                    {
                        float minLeft = float.MaxValue;
                        float maxRight = float.MinValue;
                        float totalWidth = 0;
                        foreach (var s in shapes)
                        {
                            if (s.Left < minLeft) minLeft = s.Left;
                            if (s.Left + s.Width > maxRight) maxRight = s.Left + s.Width;
                            totalWidth += s.Width;
                        }

                        float center = (minLeft + maxRight) / 2;
                        float currentLeft = center - (totalWidth / 2);

                        foreach (var s in shapes)
                        {
                            s.Left = currentLeft;
                            currentLeft += s.Width;
                        }
                    }
                    else
                    {
                        var masterShape = manager.GetReferenceShape();
                        if (masterShape == null) return false;

                        int masterIndex = shapes.FindIndex(s => s.Id == masterShape.Id);
                        if (masterIndex == -1) masterIndex = 0;

                        // Pack leftwards from master
                        for (int i = masterIndex - 1; i >= 0; i--)
                        {
                            shapes[i].Left = shapes[i + 1].Left - shapes[i].Width;
                        }

                        // Pack rightwards from master
                        for (int i = masterIndex + 1; i < shapes.Count; i++)
                        {
                            shapes[i].Left = shapes[i - 1].Left + shapes[i - 1].Width;
                        }
                    }
                }
                else
                {
                    shapes.Sort((a, b) => a.Top.CompareTo(b.Top));

                    if (packToCenter)
                    {
                        float minTop = float.MaxValue;
                        float maxBottom = float.MinValue;
                        float totalHeight = 0;
                        foreach (var s in shapes)
                        {
                            if (s.Top < minTop) minTop = s.Top;
                            if (s.Top + s.Height > maxBottom) maxBottom = s.Top + s.Height;
                            totalHeight += s.Height;
                        }

                        float center = (minTop + maxBottom) / 2;
                        float currentTop = center - (totalHeight / 2);

                        foreach (var s in shapes)
                        {
                            s.Top = currentTop;
                            currentTop += s.Height;
                        }
                    }
                    else
                    {
                        var masterShape = manager.GetReferenceShape();
                        if (masterShape == null) return false;

                        int masterIndex = shapes.FindIndex(s => s.Id == masterShape.Id);
                        if (masterIndex == -1) masterIndex = 0;

                        // Pack upwards from master
                        for (int i = masterIndex - 1; i >= 0; i--)
                        {
                            shapes[i].Top = shapes[i + 1].Top - shapes[i].Height;
                        }

                        // Pack downwards from master
                        for (int i = masterIndex + 1; i < shapes.Count; i++)
                        {
                            shapes[i].Top = shapes[i - 1].Top + shapes[i - 1].Height;
                        }
                    }
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
