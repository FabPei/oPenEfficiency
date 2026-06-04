using System;
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
        Name = "Rectify Rotation",
        Tooltip = "Rectify rotation",
        IconData = "M12,5V2L8,6L12,10V7A5,5 0 0,1 17,12C17,14.76 14.76,17 12,17C9.24,17 7,14.76 7,12H5A7,7 0 0,0 12,19C15.87,19 19,15.87 19,12C19,8.13 15.87,5 12,5Z",
        Color = "#D946EF",
        Description = "Resets the rotation of all selected shapes to 0 degrees. Perfect for fixing accidents after manual rotation.",
        DetailedHelpText = "### Rectify Rotation\nResets the rotation angle of all selected shapes to 0 degrees, straightening any accidentally rotated objects.",
        Keywords = "rotate, rectify, reset, zero, straighten, fix",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class RectifyRotationFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            var shapeRange = manager.GetSelectedShapes();
            if (shapeRange == null || shapeRange.Count == 0) return false;

            try
            {
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    
                    if (shape.Type == Office.MsoShapeType.msoLine || shape.Connector == Office.MsoTriState.msoTrue)
                    {
                        try
                        {
                            if (shape.Width > shape.Height)
                                shape.Height = 0;
                            else
                                shape.Width = 0;
                        }
                        catch (Exception ex)
                        {
                            ExceptionLogger.Log(ex, "RectifyRotationFeature.Line");
                        }
                        continue;
                    }

                    float currentRotation = shape.Rotation;
                    currentRotation = currentRotation % 360;
                    if (currentRotation < 0) currentRotation += 360;

                    float snappedRotation = (float)Math.Round(currentRotation / 90.0) * 90;
                    if (snappedRotation >= 360) snappedRotation = 0;

                    shape.Rotation = snappedRotation;
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "RectifyRotationFeature.Execute");
                return false;
            }
        }
    }
}
