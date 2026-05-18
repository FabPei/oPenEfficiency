using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Align selected shapes to the right edge of the first/master object.
    /// With a single selection, snaps the shape's right edge to the nearest vertical guide line
    /// to the right, or to the slide's right edge if no guide is present.
    /// The "move until next object" behaviour is available via the context menu.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnAlignRight",
        Name = "Align Right",
        Tooltip = "Snap to right guide / align right edges",
        IconData = "M20,2H22V22H20V2M2,10H18V14H2V10Z",
        Color = "#10B981",
        Description = "Multi-select: aligns right edges to the anchor object. Single select: snaps right edge to the nearest vertical guide line (or slide edge). Right-click for 'Move Until Object'.",
        DetailedHelpText = "### Align Right\nAligns the right edges of all selected shapes to the right edge of the first shape.",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class AlignToFirstRightFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return AlignmentHelpers.ExecuteAlignment(manager, "right");
        }
    }
}
