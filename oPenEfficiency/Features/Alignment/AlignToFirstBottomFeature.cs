using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Align selected shapes to the bottom edge of the first/master object.
    /// With a single selection, snaps the shape's bottom edge to the nearest horizontal guide line
    /// below, or to the slide's bottom edge if no guide is present.
    /// The "move until next object" behaviour is available via the context menu.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnAlignBottom",
        Name = "Align Bottom",
        Tooltip = "Snap to bottom guide / align bottom edges",
        IconData = "M22,20V22H2V20H22M10,2V18H14V2H10Z",
        Color = "#10B981",
        Description = "Multi-select: aligns bottom edges to the anchor object. Single select: snaps bottom edge to the nearest horizontal guide line (or slide edge). Right-click for 'Move Until Object'.",
        DetailedHelpText = "### Align Bottom\nAligns the bottom edges of all selected shapes to the bottom edge of the first shape in the selection.",
        Keywords = "bottom edge, snap to bottom, align bottom, horizontal guide",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class AlignToFirstBottomFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return AlignmentHelpers.ExecuteAlignment(manager, "bottom");
        }
    }
}
