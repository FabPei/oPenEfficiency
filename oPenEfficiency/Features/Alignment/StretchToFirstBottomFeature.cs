using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Stretch the height of selected shapes DOWN so their Bottom touches the Top of the first object.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnStretchBottom",
        Name = "Stretch Bottom",
        Tooltip = "Stretch bottom",
        IconData = "M13,2V14H16L12,18L8,14H11V2H13M22,22V20H2V22H22Z",
        Color = "#6366F1",
        Description = "Stretches the bottom edge of selected shapes to align with the bottom edge of the anchor.",
        DetailedHelpText = "### Stretch to Bottom\nStretches dependent shapes downward until their bottom edge aligns with the bottom edge of the master shape.",
        Keywords = "stretch down, expand bottom edge, align bottom stretch",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class StretchToFirstBottomFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var shapeRange = manager.GetSelectedShapes();
                if (shapeRange == null || shapeRange.Count < 2) return false;

                var refShape = manager.GetReferenceShape();
                float refTop = refShape.Top;
                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];
                    if (shape.Top + shape.Height < refTop) // Only if shape is above reference
                    {
                        shape.Height = refTop - shape.Top;
                    }
                }
                return true;
            }
            catch (System.Exception ex)
            {
                ExceptionLogger.Log(ex, "StretchToFirstBottomFeature.Execute");
                return false;
            }
        }
    }
}
