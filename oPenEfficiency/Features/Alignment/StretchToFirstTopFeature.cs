using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Stretch the height of selected shapes UP so their Top touches the Bottom of the first object.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnStretchTop",
        Name = "Stretch Top",
        Tooltip = "Stretch top",
        IconData = "M13,22V10H16L12,6L8,10H11V22H13M22,2V4H2V2H22Z",
        Color = "#6366F1",
        Description = "Stretches the top edge of selected shapes to align with the top edge of the anchor.",
        DetailedHelpText = "### Stretch to Top\nStretches dependent shapes upward until their top edge aligns with the top edge of the master shape.",
        Keywords = "stretch up, expand top edge, align top stretch",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class StretchToFirstTopFeature
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
                    if (shape.Top > refBottom) // Only if shape is below reference
                    {
                        float newHeight = (shape.Top + shape.Height) - refBottom;
                        shape.Top = refBottom;
                        shape.Height = newHeight;
                    }
                }
                return true;
            }
            catch (System.Exception ex)
            {
                ExceptionLogger.Log(ex, "StretchToFirstTopFeature.Execute");
                return false;
            }
        }
    }
}
