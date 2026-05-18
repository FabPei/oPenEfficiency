using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Stretch selected shapes so their Right edge matches the first object's Right edge.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnStretchRightEdge",
        Name = "Stretch Right Edge",
        Tooltip = "Stretch right edge",
        IconData = "M2,2H4V22H2V2M13,12L9,8V11H22V13H9V16L13,12Z",
        Color = "#6366F1",
        Description = "Extends the right edge of selected shapes to meet the left edge of the anchor.",
        DetailedHelpText = "### Stretch to Right Edge\nExpands each shape's right edge to meet the left edge of the shape directly to its right.",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class StretchToFirstRightEdgeFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var shapeRange = manager.GetSelectedShapes();
                if (shapeRange == null || shapeRange.Count < 2) return false;

                var refShape = manager.GetReferenceShape();
                float targetRight = refShape.Left + refShape.Width;
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    if (targetRight > shape.Left) // Ensure we don't invert
                    {
                        shape.Width = targetRight - shape.Left;
                    }
                }
                return true;
            }
            catch (System.Exception ex)
            {
                ExceptionLogger.Log(ex, "StretchToFirstRightEdgeFeature.Execute");
                return false;
            }
        }
    }
}
