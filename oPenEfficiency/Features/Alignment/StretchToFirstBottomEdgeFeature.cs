using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Stretch selected shapes so their Bottom edge matches the first object's Bottom edge.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnStretchBottomEdge",
        Name = "Stretch Bottom Edge",
        Tooltip = "Stretch bottom edge",
        IconData = "M2,2V4H22V2H2M12,13L8,9H11V2H13V9H16L12,13Z",
        Color = "#6366F1",
        Description = "Extends the bottom edge of selected shapes to meet the top edge of the anchor.",
        DetailedHelpText = "### Stretch to Bottom Edge\nStretches each selected shape's bottom edge to reach the top edge of the shape directly below it, removing the gap.",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class StretchToFirstBottomEdgeFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var shapeRange = manager.GetSelectedShapes();
                if (shapeRange == null || shapeRange.Count < 2) return false;

                var refShape = manager.GetReferenceShape();
                float targetBottom = refShape.Top + refShape.Height;
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    if (targetBottom > shape.Top) // Ensure we don't invert
                    {
                        shape.Height = targetBottom - shape.Top;
                    }
                }
                return true;
            }
            catch (System.Exception ex)
            {
                ExceptionLogger.Log(ex, "StretchToFirstBottomEdgeFeature.Execute");
                return false;
            }
        }
    }
}
