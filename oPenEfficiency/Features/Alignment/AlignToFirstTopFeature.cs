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
        Tooltip = "Align top edges or snap to the top guide",
        IconData = "M22,4V2H2V4H22M10,22V6H14V22H10Z",
        Color = "#10B981",
        Description = "Multi-select: aligns top edges to the configured first/last anchor. Single select: snaps to the nearest upper horizontal guide or slide edge.",
        DetailedHelpText = "### Align Top\n\nWith multiple selected shapes, keeps the anchor object configured in Settings (**First selected** or **Last selected**) fixed and aligns every other top edge to it. With one shape, snaps the top edge to the nearest horizontal presentation, master, or layout guide above it; the slide edge is the fallback. Right-click for explicit selection, slide, first-anchor, last-anchor, and collision modes. Each action creates a PowerPoint Undo entry.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class AlignToFirstTopFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return AlignmentHelpers.ExecuteAlignment(manager, "top");
        }

        public static bool ExecuteWithinSelection(PowerPointManager manager) => ExecuteExplicit(manager, () => AlignmentHelpers.AlignShapes(manager, "top"));
        public static bool ExecuteToSlide(PowerPointManager manager) => ExecuteExplicit(manager, () => AlignmentHelpers.AlignShapesToSlide(manager, "top"));
        public static bool ExecuteToFirstSelected(PowerPointManager manager) => ExecuteExplicit(manager, () => AlignmentHelpers.AlignToSelectedAnchor(manager, "top", false));
        public static bool ExecuteToLastSelected(PowerPointManager manager) => ExecuteExplicit(manager, () => AlignmentHelpers.AlignToSelectedAnchor(manager, "top", true));
        private static bool ExecuteExplicit(PowerPointManager manager, System.Func<bool> action) { try { manager?.GetApplication()?.StartNewUndoEntry(); } catch { /* Older PowerPoint hosts may not expose undo grouping. */ } return action(); }
    }
}
