using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Visuals
{
    [FeatureMetadata(Id = "BtnFlipHorizontal", Name = "Flip Horizontal", Tooltip = "Flip selected shapes horizontally", IconData = "M15 21l6-9l-6-9v18M9 21L3 12l6-9v18M12 2v20", Color = "#D946EF", Description = "Flips every selected shape horizontally around its own center.", DetailedHelpText = "### Flip Horizontal\n\nMirrors every selected shape from left to right around its own center. Right-click either Flip button to choose horizontal or vertical mirroring directly. Position and size stay unchanged, and the complete selection can be restored with one PowerPoint Undo step.", MinSelection = 1, RequiredType = OfficeSelectionType.Shapes)]
    public static class FlipHorizontalFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return ObjectFlipHelper.TransformShapes(manager, Office.MsoFlipCmd.msoFlipHorizontal, "FlipHorizontalFeature.Execute");
        }
    }

    [FeatureMetadata(Id = "BtnFlipVertical", Name = "Flip Vertical", Tooltip = "Flip selected shapes vertically", IconData = "M21 9l-9-6l-9 6h18M21 15l-9 6l-9-6h18M2 12h20", Color = "#D946EF", Description = "Flips every selected shape vertically around its own center.", DetailedHelpText = "### Flip Vertical\n\nMirrors every selected shape from top to bottom around its own center. Right-click either Flip button to choose horizontal or vertical mirroring directly. Position and size stay unchanged, and the complete selection can be restored with one PowerPoint Undo step.", MinSelection = 1, RequiredType = OfficeSelectionType.Shapes)]
    public static class FlipVerticalFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return ObjectFlipHelper.TransformShapes(manager, Office.MsoFlipCmd.msoFlipVertical, "FlipVerticalFeature.Execute");
        }
    }

    [FeatureMetadata(Id = "BtnRotateShape", Name = "Rotate Shapes", SearchAliases = new[] { "Rotate" }, Tooltip = "Rotate selected shapes 90° clockwise", IconData = "M12 2v3a7 7 0 0 1 7 7h3a10 10 0 0 0-10-10m7 10a7 7 0 0 1-7 7v3a10 10 0 0 0 10-10h-3M12 22v-3a7 7 0 0 1-7-7H2a10 10 0 0 0 10 10M5 12a7 7 0 0 1 7-7V2A10 10 0 0 0 2 12h3M12 2l-3 4h6l-3-4M22 12l-4-3v6l4-3M12 22l3-4H9l3 4M2 12l4 3V9l-4 3z", Color = "#D946EF", Description = "Rotates every selected shape 90° clockwise; the context menu provides 90° left and 45° steps.", DetailedHelpText = "### Rotate Shapes\n\nClick to rotate every selected shape 90° clockwise. Right-click for 90° left, 45° right, or 45° left. Each shape rotates around its own center, and the complete selection can be restored with one PowerPoint Undo step.", MinSelection = 1, RequiredType = OfficeSelectionType.Shapes)]
    public static class RotateShapeFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            // Left click is 90 right
            return ExecuteSpecific(manager, 90f);
        }

        public static bool ExecuteSpecific(PowerPointManager manager, float degrees)
        {
            if (manager == null || float.IsNaN(degrees) || float.IsInfinity(degrees)) return false;
            var shapes = manager.GetSelectedShapes();
            if (shapes == null || shapes.Count == 0) return false;

            var states = new List<RotationState>();
            try
            {
                for (int i = 1; i <= shapes.Count; i++)
                    states.Add(new RotationState(shapes[i]));

                try { manager.GetApplication()?.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "RotateShapeFeature.StartNewUndoEntry"); }

                for (int i = 1; i <= shapes.Count; i++)
                {
                    shapes[i].IncrementRotation(degrees);
                }
                return true;
            }
            catch (Exception ex)
            {
                foreach (var state in states)
                {
                    try { state.Shape.Rotation = state.Rotation; }
                    catch (Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "RotateShapeFeature.Rollback"); }
                }
                ExceptionLogger.Log(ex, "RotateShapeFeature.ExecuteSpecific");
                return false;
            }
        }

        private sealed class RotationState
        {
            public RotationState(Shape shape)
            {
                Shape = shape;
                Rotation = shape.Rotation;
            }

            public Shape Shape { get; }
            public float Rotation { get; }
        }
    }

    internal static class ObjectFlipHelper
    {
        public static bool TransformShapes(
            PowerPointManager manager,
            Office.MsoFlipCmd direction,
            string logContext)
        {
            if (manager == null) return false;
            var shapes = manager.GetSelectedShapes();
            if (shapes == null || shapes.Count == 0) return false;

            var changedShapes = new List<Shape>();
            try
            {
                try { manager.GetApplication()?.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, logContext + ".StartNewUndoEntry"); }

                for (int i = 1; i <= shapes.Count; i++)
                {
                    shapes[i].Flip(direction);
                    changedShapes.Add(shapes[i]);
                }
                return true;
            }
            catch (Exception ex)
            {
                foreach (var shape in changedShapes)
                {
                    try { shape.Flip(direction); }
                    catch (Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, logContext + ".Rollback"); }
                }
                ExceptionLogger.Log(ex, logContext);
                return false;
            }
        }
    }
}
