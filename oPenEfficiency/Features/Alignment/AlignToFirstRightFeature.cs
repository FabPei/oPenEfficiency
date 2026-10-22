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
        Tooltip = "Align right edges or snap to the right guide",
        IconData = "M20,2H22V22H20V2M2,10H18V14H2V10Z",
        Color = "#10B981",
        Description = "Multi-select: aligns right edges to the configured first/last anchor. Single select: snaps to the nearest right-side vertical guide or slide edge.",
        DetailedHelpText = "### Align Right\n\nWith multiple selected shapes, keeps the anchor object configured in Settings (**First selected** or **Last selected**) fixed and aligns every other right edge to it. With one shape, snaps the right edge to the nearest vertical presentation, master, or layout guide at or right of that edge, so the shape never moves left; the slide edge is the fallback. Right-click for explicit selection, slide, first-anchor, last-anchor, and collision modes. Each action creates a PowerPoint Undo entry.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class AlignToFirstRightFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return AlignmentHelpers.ExecuteAlignment(manager, "right");
        }

        public static bool ExecuteWithinSelection(PowerPointManager manager) => ExecuteExplicit(manager, () => AlignmentHelpers.AlignShapes(manager, "right"));
        public static bool ExecuteToSlide(PowerPointManager manager) => ExecuteExplicit(manager, () => AlignmentHelpers.AlignShapesToSlide(manager, "right"));
        public static bool ExecuteToFirstSelected(PowerPointManager manager) => ExecuteExplicit(manager, () => AlignmentHelpers.AlignToSelectedAnchor(manager, "right", false));
        public static bool ExecuteToLastSelected(PowerPointManager manager) => ExecuteExplicit(manager, () => AlignmentHelpers.AlignToSelectedAnchor(manager, "right", true));
        private static bool ExecuteExplicit(PowerPointManager manager, System.Func<bool> action) { try { manager?.GetApplication()?.StartNewUndoEntry(); } catch { /* Older PowerPoint hosts may not expose undo grouping. */ } return action(); }
    }
}
