using System;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Clones the current selection adjacently in a specified direction.
    /// </summary>
    public static class CloneSelectionFeature
    {
        public enum Direction
        {
            Left,
            Right,
            Top,
            Bottom
        }

        public static bool Execute(PowerPointManager manager, Direction dir)
        {
            if (manager == null) return false;
            var app = manager.GetApplication();
            if (app.ActiveWindow?.Selection?.Type != PpSelectionType.ppSelectionShapes) return false;

            try
            {
                var selection = app.ActiveWindow.Selection.ShapeRange;
                if (selection == null || selection.Count == 0) return false;

                // Calculate bounding box of the selection
                float minLeft = float.MaxValue, minTop = float.MaxValue;
                float maxRight = float.MinValue, maxBottom = float.MinValue;

                for (int i = 1; i <= selection.Count; i++)
                {
                    var shp = selection[i];
                    if (shp.Left < minLeft) minLeft = shp.Left;
                    if (shp.Top < minTop) minTop = shp.Top;
                    if (shp.Left + shp.Width > maxRight) maxRight = shp.Left + shp.Width;
                    if (shp.Top + shp.Height > maxBottom) maxBottom = shp.Top + shp.Height;
                }

                float totalWidth = maxRight - minLeft;
                float totalHeight = maxBottom - minTop;

                float targetLeft = minLeft;
                float targetTop = minTop;

                switch (dir)
                {
                    case Direction.Left:   targetLeft = minLeft - totalWidth; break;
                    case Direction.Right:  targetLeft = minLeft + totalWidth; break;
                    case Direction.Top:    targetTop = minTop - totalHeight; break;
                    case Direction.Bottom: targetTop = minTop + totalHeight; break;
                }

                // Duplicate the selection
                var clones = selection.Duplicate();

                // Applying Left and Top to a ShapeRange moves the entire bounding block proportionally.
                // However, if the duplicated block has a different bounding box due to PowerPoint's native offset,
                // we should explicitly calculate the shift.
                float currentCloneMinLeft = float.MaxValue;
                float currentCloneMinTop = float.MaxValue;

                for (int i = 1; i <= clones.Count; i++)
                {
                    var shp = clones[i];
                    if (shp.Left < currentCloneMinLeft) currentCloneMinLeft = shp.Left;
                    if (shp.Top < currentCloneMinTop) currentCloneMinTop = shp.Top;
                }

                float shiftX = targetLeft - currentCloneMinLeft;
                float shiftY = targetTop - currentCloneMinTop;

                // Apply the exact coordinate shift to each cloned shape to maintain perfection
                for (int i = 1; i <= clones.Count; i++)
                {
                    clones[i].Left += shiftX;
                    clones[i].Top += shiftY;
                }

                clones.Select();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CloneSelectionFeature.Execute");
                return false;
            }
        }
    }
}