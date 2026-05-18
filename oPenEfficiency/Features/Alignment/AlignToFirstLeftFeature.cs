using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Align selected shapes to the left edge of the first/master object.
    /// With a single selection, snaps the shape's left edge to the nearest vertical guide line
    /// to the left, or to the slide's left edge if no guide is present.
    /// The "move until next object" behaviour is available via the context menu.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnAlignLeft",
        Name = "Align Left",
        Tooltip = "Snap to left guide / align left edges",
        IconData = "M4,2H2V22H4V2M22,10H6V14H22V10Z",
        Color = "#10B981",
        Description = "Multi-select: aligns left edges to the anchor object. Single select: snaps left edge to the nearest vertical guide line (or slide edge). Right-click for 'Move Until Object'.",
        DetailedHelpText = "### Align Left\nAligns the left edges of all selected shapes to the left edge of the first shape. Ideal for snapping columns into a clean left margin.",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class AlignToFirstLeftFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return AlignmentHelpers.ExecuteAlignment(manager, "left");
        }
    }
}
