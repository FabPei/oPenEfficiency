using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Stack selected shapes to the left of the first object.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnDockLeft",
        Name = "Stack Left",
        Tooltip = "Stack left",
        IconData = "M10,4V20H6V4H10M12,2H4V22H12V2M14,4H20V20H14V4Z",
        Color = "#F59E0B",
        Description = "Moves selected shapes to the left side of the anchor shape, stacking them horizontally without overlapping.",
        DetailedHelpText = "### Stack Left\nMoves dependent shapes leftward until their right edge touches the left edge of the master shape.",
        Keywords = "stack horizontally, dock left, zero gap horizontal, prepend left",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class DockToFirstLeftFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var shapeRange = manager.GetSelectedShapes();
                if (shapeRange == null || shapeRange.Count < 2) return false;

                var refShape = manager.GetReferenceShape();
                float currentLeft = refShape.Left;
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    if (shape.Id == refShape.Id) continue;
                    shape.Left = currentLeft - shape.Width;
                    currentLeft = shape.Left;
                }
                return true;
            }
            catch (System.Exception ex)
            {
                ExceptionLogger.Log(ex, "DockToFirstLeftFeature.Execute");
                return false;
            }
        }
    }
}
