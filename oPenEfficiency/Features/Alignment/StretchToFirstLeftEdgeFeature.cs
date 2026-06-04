using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Stretch selected shapes so their Left edge matches the first object's Left edge.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnStretchLeftEdge",
        Name = "Stretch Left Edge",
        Tooltip = "Stretch left edge",
        IconData = "M20,2H22V22H20V2M11,12L15,8V11H2V13H15V16L11,12Z",
        Color = "#6366F1",
        Description = "Extends the left edge of selected shapes to meet the right edge of the anchor.",
        DetailedHelpText = "### Stretch to Left Edge\nExpands each shape's left edge to meet the right edge of the shape directly to its left.",
        Keywords = "stretch to left, fill gap left, expand leftwards, touch left",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class StretchToFirstLeftEdgeFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var shapeRange = manager.GetSelectedShapes();
                if (shapeRange == null || shapeRange.Count < 2) return false;

                var refShape = manager.GetReferenceShape();
                float targetLeft = refShape.Left;
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    float currentRight = shape.Left + shape.Width;
                    if (currentRight > targetLeft) // Ensure we don't invert
                    {
                        shape.Left = targetLeft;
                        shape.Width = currentRight - targetLeft;
                    }
                }
                return true;
            }
            catch (System.Exception ex)
            {
                ExceptionLogger.Log(ex, "StretchToFirstLeftEdgeFeature.Execute");
                return false;
            }
        }
    }
}
