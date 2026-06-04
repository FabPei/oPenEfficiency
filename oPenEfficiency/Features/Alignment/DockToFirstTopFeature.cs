using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Stack selected shapes on top of the first object.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnDockTop",
        Name = "Stack Top",
        Tooltip = "Stack top",
        IconData = "M4,10V6H20V10H4M2,12V4H22V12H2M4,14V20H20V14H4Z",
        Color = "#F59E0B",
        Description = "Moves selected shapes above the anchor shape, stacking them vertically.",
        DetailedHelpText = "### Stack Top\nMoves dependent shapes upward until their bottom edge touches the top edge of the master shape.",
        Keywords = "stack vertically, dock above, zero gap vertical, prepend top",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class DockToFirstTopFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var shapeRange = manager.GetSelectedShapes();
                if (shapeRange == null || shapeRange.Count < 2) return false;

                var refShape = manager.GetReferenceShape();
                float currentTop = refShape.Top;
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    if (shape.Id == refShape.Id) continue;
                    shape.Top = currentTop - shape.Height;
                    currentTop = shape.Top;
                }
                return true;
            }
            catch (System.Exception ex)
            {
                ExceptionLogger.Log(ex, "DockToFirstTopFeature.Execute");
                return false;
            }
        }
    }
}
