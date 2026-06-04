using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Stretch the width of selected shapes RIGHT so their Right touches the Left of the first object.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnStretchRight",
        Name = "Stretch Right",
        Tooltip = "Stretch right",
        IconData = "M2,13H14V16L18,12L14,8V11H2V13M22,2H20V22H22V2Z",
        Color = "#6366F1",
        Description = "Stretches the right edge of selected shapes to align with the right edge of the anchor.",
        DetailedHelpText = "### Stretch to Right\nExpands dependent shapes rightward to align their right edge with the master shape's right edge.",
        Keywords = "stretch right, expand right edge, align right stretch",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class StretchToFirstRightFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var shapeRange = manager.GetSelectedShapes();
                if (shapeRange == null || shapeRange.Count < 2) return false;

                var refShape = manager.GetReferenceShape();
                float refLeft = refShape.Left;
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    if (shape.Left + shape.Width < refLeft) // Only if shape is to the left of reference
                    {
                        shape.Width = refLeft - shape.Left;
                    }
                }
                return true;
            }
            catch (System.Exception ex)
            {
                ExceptionLogger.Log(ex, "StretchToFirstRightFeature.Execute");
                return false;
            }
        }
    }
}
