using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Stack selected shapes to the right of the first object.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnDockRight",
        Name = "Stack Right",
        Tooltip = "Stack right",
        IconData = "M14,4V20H18V4H14M12,2H20V22H12V2M10,4H4V20H10V4Z",
        Color = "#F59E0B",
        Description = "Moves selected shapes to the right side of the anchor shape, stacking them horizontally.",
        DetailedHelpText = "### Stack Right\nMoves dependent shapes rightward until their left edge touches the right edge of the master shape.",
        Keywords = "stack horizontally, dock right, zero gap horizontal, append right",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class DockToFirstRightFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var shapeRange = manager.GetSelectedShapes();
                if (shapeRange == null || shapeRange.Count < 2) return false;

                var refShape = manager.GetReferenceShape();
                float currentRight = refShape.Left + refShape.Width;
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    if (shape.Id == refShape.Id) continue;
                    shape.Left = currentRight;
                    currentRight += shape.Width;
                }
                return true;
            }
            catch (System.Exception ex)
            {
                ExceptionLogger.Log(ex, "DockToFirstRightFeature.Execute");
                return false;
            }
        }
    }
}
