using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Stretch selected shapes so their Top edge matches the first object's Top edge.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnStretchTopEdge",
        Name = "Stretch Top Edge",
        Tooltip = "Stretch top edge",
        IconData = "M2,20V22H22V20H2M12,11L8,15H11V22H13V15H16L12,11Z",
        Color = "#6366F1",
        Description = "Extends the top edge of selected shapes to meet the bottom edge of the anchor.",
        DetailedHelpText = "### Stretch to Top Edge\nStretches each selected shape's top edge to reach the bottom edge of the shape directly above it.",
        Keywords = "stretch to top, fill gap up, expand upwards, touch top",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class StretchToFirstTopEdgeFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var shapeRange = manager.GetSelectedShapes();
                if (shapeRange == null || shapeRange.Count < 2) return false;

                var refShape = manager.GetReferenceShape();
                float targetTop = refShape.Top;
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    float currentBottom = shape.Top + shape.Height;
                    if (currentBottom > targetTop) // Ensure we don't invert
                    {
                        shape.Top = targetTop;
                        shape.Height = currentBottom - targetTop;
                    }
                }
                return true;
            }
            catch (System.Exception ex)
            {
                ExceptionLogger.Log(ex, "StretchToFirstTopEdgeFeature.Execute");
                return false;
            }
        }
    }
}
