using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Snaps the rotation of all selected shapes to the nearest 90-degree increment (0, 90, 180, 270).
    /// </summary>
    [FeatureMetadata(
        Id = "BtnRectifyRotation",
        Name = "Snap Rotation to 90°",
        SearchAliases = new[] { "Rectify Rotation" },
        Tooltip = "Snap rotation to the nearest 90°",
        IconData = "M12,5V2L8,6L12,10V7A5,5 0 0,1 17,12C17,14.76 14.76,17 12,17C9.24,17 7,14.76 7,12H5A7,7 0 0,0 12,19C15.87,19 19,15.87 19,12C19,8.13 15.87,5 12,5Z",
        Color = "#D946EF",
        Description = "Snaps every selected shape to the nearest 90° orientation and straightens near-horizontal or near-vertical lines.",
        DetailedHelpText = "### Snap Rotation to 90°\n\nSnaps each selected shape to its nearest 0°, 90°, 180°, or 270° orientation. Lines and connectors are straightened to the nearer horizontal or vertical axis. Right-click to set the complete selection explicitly to 0°, 90°, 180°, or 270°. Each operation can be restored with one PowerPoint Undo step.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class RectifyRotationFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            var shapeRange = manager.GetSelectedShapes();
            if (shapeRange == null || shapeRange.Count == 0) return false;

            var states = new List<ShapeState>();
            try
            {
                for (int i = 1; i <= shapeRange.Count; i++)
                    states.Add(new ShapeState(shapeRange[i]));

                try { manager.GetApplication()?.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "RectifyRotationFeature.StartNewUndoEntry"); }

                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    
                    if (shape.Type == Office.MsoShapeType.msoLine || shape.Connector == Office.MsoTriState.msoTrue)
                    {
                        if (shape.Width > shape.Height)
                            shape.Height = 0;
                        else
                            shape.Width = 0;
                        continue;
                    }

                    float currentRotation = shape.Rotation;
                    currentRotation = currentRotation % 360;
                    if (currentRotation < 0) currentRotation += 360;

                    float snappedRotation = (float)Math.Round(
                        currentRotation / 90.0,
                        MidpointRounding.AwayFromZero) * 90;
                    if (snappedRotation >= 360) snappedRotation = 0;

                    shape.Rotation = snappedRotation;
                }
                return true;
            }
            catch (Exception ex)
            {
                foreach (var state in states)
                {
                    try { state.Restore(); }
                    catch (Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "RectifyRotationFeature.Rollback"); }
                }
                ExceptionLogger.Log(ex, "RectifyRotationFeature.Execute");
                return false;
            }
        }

        public static bool ExecuteSpecific(PowerPointManager manager, float rotation)
        {
            if (manager == null || float.IsNaN(rotation) || float.IsInfinity(rotation)) return false;
            var shapeRange = manager.GetSelectedShapes();
            if (shapeRange == null || shapeRange.Count == 0) return false;

            float normalized = rotation % 360f;
            if (normalized < 0f) normalized += 360f;
            var states = new List<ShapeState>();
            try
            {
                for (int i = 1; i <= shapeRange.Count; i++)
                    states.Add(new ShapeState(shapeRange[i]));

                try { manager.GetApplication()?.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "RectifyRotationFeature.ExecuteSpecific.StartNewUndoEntry"); }

                for (int i = 1; i <= shapeRange.Count; i++)
                    shapeRange[i].Rotation = normalized;
                return true;
            }
            catch (Exception ex)
            {
                foreach (var state in states)
                {
                    try { state.Restore(); }
                    catch (Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "RectifyRotationFeature.ExecuteSpecific.Rollback"); }
                }
                ExceptionLogger.Log(ex, "RectifyRotationFeature.ExecuteSpecific");
                return false;
            }
        }

        private sealed class ShapeState
        {
            public ShapeState(Shape shape)
            {
                Shape = shape;
                Left = shape.Left;
                Top = shape.Top;
                Width = shape.Width;
                Height = shape.Height;
                Rotation = shape.Rotation;
            }

            public Shape Shape { get; }
            public float Left { get; }
            public float Top { get; }
            public float Width { get; }
            public float Height { get; }
            public float Rotation { get; }

            public void Restore()
            {
                Shape.Rotation = Rotation;
                Shape.Width = Width;
                Shape.Height = Height;
                Shape.Left = Left;
                Shape.Top = Top;
            }
        }
    }
}
