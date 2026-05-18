using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Align selected shapes to the top edge of the first/master object.
    /// With a single selection, snaps the shape's top edge to the nearest horizontal guide line
    /// above, or to the slide's top edge if no guide is present.
    /// The "move until next object" behaviour is available via the context menu.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnAlignTop",
        Name = "Align Top",
        Tooltip = "Snap to top guide / align top edges",
        IconData = "M22,4V2H2V4H22M10,22V6H14V22H10Z",
        Color = "#10B981",
        Description = "Multi-select: aligns top edges to the anchor object. Single select: snaps top edge to the nearest horizontal guide line (or slide edge). Right-click for 'Move Until Object'.",
        DetailedHelpText = "### Align Top\nAligns the top edges of all selected shapes to the top edge of the first shape.",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class AlignToFirstTopFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return AlignmentHelpers.ExecuteAlignment(manager, "top");
        }
    }
}
