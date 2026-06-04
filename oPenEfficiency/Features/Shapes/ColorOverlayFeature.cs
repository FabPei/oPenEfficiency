using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.Features.Shapes
{
    [FeatureMetadata(
        Id = "BtnColorOverlay",
        Name = "Color Overlay",
        Tooltip = "Cut bottom shape with top shape and color the overlay",
        IconData = "M12,2A10,10 0 0,0 2,12A10,10 0 0,0 12,22A10,10 0 0,0 22,12A10,10 0 0,0 12,2M12,4A8,8 0 0,1 20,12A8,8 0 0,1 12,20A8,8 0 0,1 4,12A8,8 0 0,1 12,4M12,6A6,6 0 0,0 6,12A6,6 0 0,0 12,18A6,6 0 0,0 18,12A6,6 0 0,0 12,6Z",
        Color = "#8E24AA", // Purple
        Description = "Applies color from the second shape to the overlapping area of the first shape.",
        DetailedHelpText = "This feature combines two selected shapes. The first selected shape (bottom) keeps its original color where it does not overlap, but takes the color of the second selected shape (top) where they overlap. The two resulting pieces are then grouped together.",
        Keywords = "intersect shapes color, boolean color merge, overlay fill, subtract color",
        MinSelection = 2,
        MaxSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class ColorOverlayFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try
            {
                var shapes = manager.GetSelectedShapes();
                if (shapes == null || shapes.Count != 2)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Please select exactly 2 shapes.", "Info",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return false;
                }

                var app = manager.GetApplication();
                var slide = (Slide)app.ActiveWindow.View.Slide;
                
                var shape1 = shapes[1]; // Bottom / Base shape
                var shape2 = shapes[2]; // Top / Overlay shape

                // Extract colors
                bool s1HasFill = shape1.Fill.Visible == Office.MsoTriState.msoTrue;
                int s1Color = s1HasFill ? shape1.Fill.ForeColor.RGB : 0xFFFFFF;

                bool s2HasFill = shape2.Fill.Visible == Office.MsoTriState.msoTrue;
                int s2Color = s2HasFill ? shape2.Fill.ForeColor.RGB : 0x000000;

                // Save positions
                float orig1L = shape1.Left;
                float orig1T = shape1.Top;
                float orig2L = shape2.Left;
                float orig2T = shape2.Top;

                // Part 1: Intersection
                var s1_copy1 = shape1.Duplicate()[1];
                s1_copy1.Left = orig1L; s1_copy1.Top = orig1T;

                var s2_copy1 = shape2.Duplicate()[1];
                s2_copy1.Left = orig2L; s2_copy1.Top = orig2T;

                s1_copy1.Select(Office.MsoTriState.msoTrue);
                s2_copy1.Select(Office.MsoTriState.msoFalse);
                
                app.CommandBars.ExecuteMso("ShapesIntersect");
                
                Shape intersectShape = null;
                if (app.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionShapes && app.ActiveWindow.Selection.ShapeRange.Count > 0)
                {
                    intersectShape = app.ActiveWindow.Selection.ShapeRange[1];
                    if (s2HasFill)
                    {
                        intersectShape.Fill.Visible = Office.MsoTriState.msoTrue;
                        intersectShape.Fill.ForeColor.RGB = s2Color;
                    }
                    else
                    {
                        intersectShape.Fill.Visible = Office.MsoTriState.msoFalse;
                    }
                }

                // Part 2: Subtraction
                var s1_copy2 = shape1.Duplicate()[1];
                s1_copy2.Left = orig1L; s1_copy2.Top = orig1T;

                var s2_copy2 = shape2.Duplicate()[1];
                s2_copy2.Left = orig2L; s2_copy2.Top = orig2T;

                s1_copy2.Select(Office.MsoTriState.msoTrue);
                s2_copy2.Select(Office.MsoTriState.msoFalse);

                app.CommandBars.ExecuteMso("ShapesSubtract");

                Shape subtractShape = null;
                if (app.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionShapes && app.ActiveWindow.Selection.ShapeRange.Count > 0)
                {
                    subtractShape = app.ActiveWindow.Selection.ShapeRange[1];
                    // Should retain s1Color, but let's be explicit
                    if (s1HasFill)
                    {
                        subtractShape.Fill.Visible = Office.MsoTriState.msoTrue;
                        subtractShape.Fill.ForeColor.RGB = s1Color;
                    }
                    else
                    {
                        subtractShape.Fill.Visible = Office.MsoTriState.msoFalse;
                    }
                }

                // Group them if both exist
                var finalShapes = new List<string>();
                if (subtractShape != null) finalShapes.Add(subtractShape.Name);
                if (intersectShape != null) finalShapes.Add(intersectShape.Name);

                if (finalShapes.Count > 1)
                {
                    var group = slide.Shapes.Range(finalShapes.ToArray()).Group();
                    group.Select(Office.MsoTriState.msoTrue);
                }
                else if (finalShapes.Count == 1)
                {
                    slide.Shapes[finalShapes[0]].Select(Office.MsoTriState.msoTrue);
                }

                // Delete original shapes
                shape1.Delete();
                shape2.Delete();

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ColorOverlayFeature.Execute", showErrorUI: true);
                return false;
            }
        }
    }
}
