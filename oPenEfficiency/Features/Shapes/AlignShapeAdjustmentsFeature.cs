using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using oPenEfficiency.Models;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Features.Shapes
{
    [FeatureMetadata(
        Id = "BtnAlignShapeAdjustments",
        Name = "Shape Transformation / Vertex Align",
        Tooltip = "Align Shape Proportions",
        IconData = "M12,2L4.5,20.29L5.21,21L12,18L18.79,21L19.5,20.29L12,2M12,5.83L17.5,19L12,16.62L6.5,19L12,5.83Z", // Compass/Transform icon
        Color = "#FBBF24",
        MinSelection = 2,
        RequiredType = PowerPoint.PpSelectionType.ppSelectionShapes,
        Description = "Reads the exact proportions (like shaft thickness vs. head width of a Block Arrow or Pentagon) of a Master shape and forces all other selected shapes to adopt identical vertex ratios.",
        DetailedHelpText = "### Shape Transformation / Vertex Align\n\nTransfers vertex-level adjustment values (yellow diamond handles) from the first shape to all other selected shapes of the same type, normalizing arrowhead sizes, corner radii, and tip angles."
    )]
    public static class AlignShapeAdjustmentsFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try
            {
                var shapeRange = manager.GetSelectedShapes();
                if (shapeRange == null || shapeRange.Count < 2) return false;

                var refShape = manager.GetReferenceShape();

                // If the reference shape does not have adjustments, nothing to do
                if (refShape.Adjustments == null || refShape.Adjustments.Count == 0)
                {
                    return false;
                }

                bool modified = false;

                for (int i = 1; i <= shapeRange.Count; i++)
                {
                    var shape = shapeRange[i];

                    // Don't apply to the reference shape itself
                    if (shape.Id == refShape.Id) continue;

                    // Only apply if they are the exact same shape geometry
                    if (shape.Type == refShape.Type && shape.AutoShapeType == refShape.AutoShapeType)
                    {
                        if (shape.Adjustments != null && shape.Adjustments.Count == refShape.Adjustments.Count)
                        {
                            for (int adjIdx = 1; adjIdx <= refShape.Adjustments.Count; adjIdx++)
                            {
                                shape.Adjustments[adjIdx] = refShape.Adjustments[adjIdx];
                            }
                            modified = true;
                        }
                    }
                }

                return modified;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex.Message, "AlignShapeAdjustmentsFeature");
                return false;
            }
        }
    }
}
