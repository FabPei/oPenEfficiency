using System;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnSnapShape",
        Name = "Snap Shape to Object",
        Tooltip = "Snap the first shape to the second",
        IconData = "M3,3H10V10H3V3M14,14H21V21H14V14M10,6H14V8H10V6M16,10V14H14V10H16Z",
        Color = "#FBBF24",
        Description = "Moves the first of exactly two selected shapes flush to the nearest side of the second shape.",
        DetailedHelpText = "### Snap Shape to Object\nSelect exactly two shapes. The first selected shape moves flush to the nearest side of the second and is centered on the perpendicular axis; size and formatting are preserved. Right-click for an explicit left, right, above, or below direction.",
        MinSelection = 2,
        MaxSelection = 2,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class SnapShapeFeature
    {
        public enum SnapDirection { Nearest, Left, Right, Above, Below }

        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, SnapDirection.Nearest);
        }

        public static bool Execute(PowerPointManager manager, SnapDirection direction)
        {
            if (manager == null) return false;
            try
            {
                var shapes = manager.GetSelectedShapes();
                if (shapes == null || shapes.Count != 2) return false;
                var moving = shapes[1];
                var target = shapes[2];
                var policyDirection = oPenEfficiency.Core.Services.ShapeSnapPlacementPolicy.SnapDirection.Nearest;
                if (direction == SnapDirection.Nearest)
                {
                    policyDirection = oPenEfficiency.Core.Services.ShapeSnapPlacementPolicy.GetNearestDirection(
                        moving.Left, moving.Top, moving.Width, moving.Height,
                        target.Left, target.Top, target.Width, target.Height);
                }
                else
                {
                    policyDirection = direction == SnapDirection.Left ? oPenEfficiency.Core.Services.ShapeSnapPlacementPolicy.SnapDirection.Left
                        : direction == SnapDirection.Right ? oPenEfficiency.Core.Services.ShapeSnapPlacementPolicy.SnapDirection.Right
                        : direction == SnapDirection.Above ? oPenEfficiency.Core.Services.ShapeSnapPlacementPolicy.SnapDirection.Above
                        : oPenEfficiency.Core.Services.ShapeSnapPlacementPolicy.SnapDirection.Below;
                }
                try { manager.GetApplication()?.StartNewUndoEntry(); } catch { /* Snapping remains valid when undo grouping is unavailable. */ }

                oPenEfficiency.Core.Services.ShapeSnapPlacementPolicy.Place(
                    policyDirection,
                    target.Left, target.Top, target.Width, target.Height,
                    moving.Width, moving.Height,
                    out float newLeft, out float newTop);
                moving.Left = newLeft;
                moving.Top = newTop;
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SnapShapeFeature.Execute");
                return false;
            }
        }
    }
}
