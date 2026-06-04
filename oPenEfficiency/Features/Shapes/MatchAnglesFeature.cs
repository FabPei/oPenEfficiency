using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Shapes
{
    [FeatureMetadata(
        Id = "BtnMatchAngles",
        Name = "Match Shape Geometry",
        Tooltip = "Aligns the yellow-diamond adjustments (e.g. arrowhead size, corner radius) of all selected shapes to match the first shape.",
        IconData = "M2,2H8V4H4V8H2V2M22,2V8H20V4H16V2H22M2,22V16H4V20H8V22H2M22,22H16V20H20V16H22V22M15,10V14L18,12L15,10M9,10L6,12L9,14V10Z",
        Color = "#8B5CF6", // Distinct Purple 
        Description = "Transfers exact vertex angles, shaft thicknesses, and corner radii across similar shapes.",
        DetailedHelpText = "### Match Shape Geometry\n\nPowerPoint warps shapes when you resize them — arrowheads become too big, pentagon tips become inconsistent, and rounded corners look different depending on size.\n\n**How it works:**\nSelect multiple shapes of the same type. The **first shape** in the selection is the Master. All other shapes will have their hidden `Adjustments` properties overwritten to match the Master's exact geometry — without changing their size or position.\n\n**Supported shapes:**\n- Block Arrows (normalizes shaft thickness and arrowhead size)\n- Chevrons & Pentagons (tip angle synchronization)\n- Rounded Rectangles (corner radius equalization)\n- Speech Bubbles / Callouts (tail width and angle)\n- Donut / Arc shapes (inner radius depth)\n\n*Note: Shapes must be the same AutoShapeType. Mixing an Arrow with a Pentagon will be skipped automatically.*",
        Keywords = "clone geometry, equal angles, identical tips, match adjustments",
        MinSelection = 2)]
    public static class MatchAnglesFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;

            try
            {
                var app = manager.GetApplication();
                
                if (app.ActiveWindow.Selection.Type != PpSelectionType.ppSelectionShapes)
                {
                    System.Windows.Forms.MessageBox.Show("Please select at least two shapes.", "Invalid Selection", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return false;
                }

                Microsoft.Office.Interop.PowerPoint.ShapeRange selection = app.ActiveWindow.Selection.ShapeRange;

                if (selection.Count < 2)
                {
                    System.Windows.Forms.MessageBox.Show("Please select at least two shapes to match geometries.", "Invalid Selection", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return false;
                }

                // Reference shape is the first one in the selection array
                Microsoft.Office.Interop.PowerPoint.Shape masterShape = selection[1];

                if (masterShape.Adjustments.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show("The primary shape has no adjustable geometry (no yellow diamonds). Select an adjustable shape (like an Arrow, Rounded Rectangle, or Pentagon) first.", "No Adjustments Available", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return false;
                }

                int modifiedCount = 0;

                // Iterate over the rest of the shapes
                for (int i = 2; i <= selection.Count; i++)
                {
                    Microsoft.Office.Interop.PowerPoint.Shape targetShape = selection[i];

                    // Safest execution: Only transfer geometry if they are the exact same shape type
                    // E.g., transferring Arrowhead properties to a Rounded Rectangle will crash PowerPoint
                    if (targetShape.AutoShapeType == masterShape.AutoShapeType && targetShape.Adjustments.Count > 0)
                    {
                        for (int a = 1; a <= masterShape.Adjustments.Count; a++)
                        {
                            try
                            {
                                // PowerPoint Adjustments are indexed strictly 1-based.
                                targetShape.Adjustments[a] = masterShape.Adjustments[a];
                            }
                            catch
                            {
                                // Some adjustments might be read-only under very specific geometric bounds.
                                // We swallow individual adjustment errors to let the outer loop continue syncing.
                            }
                        }
                        modifiedCount++;
                    }
                }

                if (modifiedCount == 0)
                {
                    System.Windows.Forms.MessageBox.Show($"Could not match geometry. The other shapes must be of EXACTLY the same shape type as the Master shape ({masterShape.AutoShapeType}).", "Geometry Mismatch", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "MatchAnglesFeature.Execute");
                System.Windows.Forms.MessageBox.Show("An error occurred while matching shape geometries.", "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }
    }
}
