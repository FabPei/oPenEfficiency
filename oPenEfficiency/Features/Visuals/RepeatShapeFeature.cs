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
        IconData = "M1,4H8V11H1Z M3,6V9H6V6Z M8.5,4H15.5V11H8.5Z M10.5,6V9H13.5V6Z M16,4H23V11H16Z M18,6V9H21V6Z M2,17H18.6L16.3,14.7L17.7,13.3L22.4,18L17.7,22.7L16.3,21.3L18.6,19H2Z",
        Color = "#F43F5E",
        Description = "Clones the selected shape multiple times to create long sequences or grids. Default direction is Right.",
        DetailedHelpText = "### Repeat Shape\nDuplicates the selected shape a configurable number of times with automatic offset spacing, instantly generating a row or column of identical shapes.",
        MinSelection = 2,
        MaxSelection = 2,
        RequiredType = OfficeSelectionType.Shapes)]
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

                float gapX = oPenEfficiency.Core.Services.RepeatShapeGeometryPlanner.InferGapX(shape1.Left, shape1.Width, shape2.Left, shape2.Width);
                float gapY = oPenEfficiency.Core.Services.RepeatShapeGeometryPlanner.InferGapY(shape1.Top, shape1.Height, shape2.Top, shape2.Height);

                // Duplicate shape2 precisely
                shape2.Select(Office.MsoTriState.msoTrue);
                var newShapeRange = shape2.Duplicate();
                var newShape = newShapeRange[1];

                var placementDirection = direction == RepeatDirection.Right ? oPenEfficiency.Core.Services.RepeatShapeGeometryPlanner.RepeatPlacementDirection.Right
                    : direction == RepeatDirection.Left ? oPenEfficiency.Core.Services.RepeatShapeGeometryPlanner.RepeatPlacementDirection.Left
                    : direction == RepeatDirection.Down ? oPenEfficiency.Core.Services.RepeatShapeGeometryPlanner.RepeatPlacementDirection.Down
                    : oPenEfficiency.Core.Services.RepeatShapeGeometryPlanner.RepeatPlacementDirection.Up;
                oPenEfficiency.Core.Services.RepeatShapeGeometryPlanner.NextPlacement(
                    placementDirection,
                    shape2.Left, shape2.Top, shape2.Width, shape2.Height,
                    newShape.Width, newShape.Height, gapX, gapY,
                    out float newLeft, out float newTop);
                newShape.Left = newLeft;
                newShape.Top = newTop;

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
