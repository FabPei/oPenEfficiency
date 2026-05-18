using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Stretch the width of selected shapes LEFT so their Left touches the Right of the first object.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnStretchLeft",
        Name = "Stretch Left",
        Tooltip = "Stretch left",
        IconData = "M22,13H10V16L6,12L10,8V11H22V13M2,2H4V22H2V2Z",
        Color = "#6366F1",
        Description = "Stretches the left edge of selected shapes to align with the left edge of the anchor.",
        DetailedHelpText = "### Stretch to Left\nExpands dependent shapes leftward to align their left edge with the master shape's left edge.",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class StretchToFirstLeftFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var shapeRange = manager.GetSelectedShapes();
                if (shapeRange == null || shapeRange.Count < 2) return false;

                var refShape = manager.GetReferenceShape();
                float refRight = refShape.Left + refShape.Width;
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    if (shape.Left > refRight) // Only if shape is to the right of reference
                    {
                        float newWidth = (shape.Left + shape.Width) - refRight;
                        shape.Left = refRight;
                        shape.Width = newWidth;
                    }
                }
                return true;
            }
            catch (System.Exception ex)
            {
                ExceptionLogger.Log(ex, "StretchToFirstLeftFeature.Execute");
                return false;
            }
        }
    }
}
