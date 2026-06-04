using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnRepeatShape",
        Name = "Repeat Shape",
        Tooltip = "Repeat shape",
        IconData = "M17,17H7V14L3,18L7,22V19H19V13H17M7,7H17V10L21,6L17,2V5H5V11H7V7Z",
        Color = "#F43F5E",
        Description = "Clones the selected shape multiple times to create long sequences or grids. Default direction is Right.",
        DetailedHelpText = "### Repeat Shape\nDuplicates the selected shape a configurable number of times with automatic offset spacing, instantly generating a row or column of identical shapes.",
        Keywords = "repeat, clone, sequence, grid, array, duplicate",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class RepeatShapeFeature
    {
        public enum RepeatDirection { Right, Left, Up, Down }

        /// <summary>
        /// Wrapper for auto-discovery - repeats shape to the right (default direction).
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, RepeatDirection.Right);
        }

        public static bool Execute(PowerPointManager manager, RepeatDirection direction)
        {
            try
            {
                var selection = manager.GetSelectedShapes();
                if (selection == null || selection.Count != 2) return false;

                var shape1 = selection[1];
                var shape2 = selection[2];

                try
                {
                    manager.GetApplication().StartNewUndoEntry();
                } catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"RepeatShapeFeature.StartNewUndoEntry error: {ex.Message}"); }

                float right1 = shape1.Left + shape1.Width;
                float bottom1 = shape1.Top + shape1.Height;
                float right2 = shape2.Left + shape2.Width;
                float bottom2 = shape2.Top + shape2.Height;

                float gapX = 0;
                if (System.Math.Max(shape1.Left, shape2.Left) > System.Math.Min(right1, right2))
                    gapX = System.Math.Max(shape1.Left, shape2.Left) - System.Math.Min(right1, right2);
                else
                    gapX = System.Math.Abs(shape1.Left - shape2.Left);

                float gapY = 0;
                if (System.Math.Max(shape1.Top, shape2.Top) > System.Math.Min(bottom1, bottom2))
                    gapY = System.Math.Max(shape1.Top, shape2.Top) - System.Math.Min(bottom1, bottom2);
                else
                    gapY = System.Math.Abs(shape1.Top - shape2.Top);

                // Duplicate shape2 precisely
                shape2.Select(Office.MsoTriState.msoTrue);
                var newShapeRange = shape2.Duplicate();
                var newShape = newShapeRange[1];

                if (direction == RepeatDirection.Right)
                {
                    newShape.Left = shape2.Left + shape2.Width + gapX;
                    newShape.Top = shape2.Top;
                }
                else if (direction == RepeatDirection.Left)
                {
                    newShape.Left = shape2.Left - newShape.Width - gapX;
                    newShape.Top = shape2.Top;
                }
                else if (direction == RepeatDirection.Down)
                {
                    newShape.Top = shape2.Top + shape2.Height + gapY;
                    newShape.Left = shape2.Left;
                }
                else if (direction == RepeatDirection.Up)
                {
                    newShape.Top = shape2.Top - newShape.Height - gapY;
                    newShape.Left = shape2.Left;
                }

                shape2.Select(Office.MsoTriState.msoTrue);
                newShape.Select(Office.MsoTriState.msoFalse);
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
