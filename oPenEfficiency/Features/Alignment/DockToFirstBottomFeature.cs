using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Stack selected shapes below the first object.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnDockBottom",
        Name = "Stack Bottom",
        Tooltip = "Stack bottom",
        IconData = "M4,14V18H20V14H4M2,12V20H22V12H2M4,10V4H20V10H4Z",
        Color = "#F59E0B",
        Description = "Moves selected shapes below the anchor shape, stacking them vertically.",
        DetailedHelpText = "### Stack Bottom\nMoves dependent shapes downward until their top edge touches the bottom edge of the master shape (zero gap docking).",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class DockToFirstBottomFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var shapeRange = manager.GetSelectedShapes();
                if (shapeRange == null || shapeRange.Count < 2) return false;

                var refShape = manager.GetReferenceShape();
                float refBottom = refShape.Top + refShape.Height;
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    if (shape.Id == refShape.Id) continue;
                    shape.Top = refBottom;
                    refBottom += shape.Height;
                }
                return true;
            }
            catch (System.Exception ex)
            {
                ExceptionLogger.Log(ex, "DockToFirstBottomFeature.Execute");
                return false;
            }
        }
    }
}
