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
        Tooltip = "Align left edges or snap to the left guide",
        IconData = "M4,2H2V22H4V2M22,10H6V14H22V10Z",
        Color = "#10B981",
        Description = "Multi-select: aligns left edges to the configured first/last anchor. Single select: snaps to the nearest left-side vertical guide or slide edge.",
        DetailedHelpText = "### Align Left\n\nWith multiple selected shapes, keeps the anchor object configured in Settings (**First selected** or **Last selected**) fixed and aligns every other left edge to it. With one shape, snaps the left edge to the nearest vertical presentation, master, or layout guide at or left of that edge, so the shape never moves right; the slide edge is the fallback. Right-click for explicit selection, slide, first-anchor, last-anchor, and collision modes. Each action creates a PowerPoint Undo entry.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class AlignToFirstLeftFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return AlignmentHelpers.ExecuteAlignment(manager, "left");
        }

        public static bool ExecuteWithinSelection(PowerPointManager manager) => ExecuteExplicit(manager, () => AlignmentHelpers.AlignShapes(manager, "left"));
        public static bool ExecuteToSlide(PowerPointManager manager) => ExecuteExplicit(manager, () => AlignmentHelpers.AlignShapesToSlide(manager, "left"));
        public static bool ExecuteToFirstSelected(PowerPointManager manager) => ExecuteExplicit(manager, () => AlignmentHelpers.AlignToSelectedAnchor(manager, "left", false));
        public static bool ExecuteToLastSelected(PowerPointManager manager) => ExecuteExplicit(manager, () => AlignmentHelpers.AlignToSelectedAnchor(manager, "left", true));
        private static bool ExecuteExplicit(PowerPointManager manager, System.Func<bool> action) { try { manager?.GetApplication()?.StartNewUndoEntry(); } catch { /* Older PowerPoint hosts may not expose undo grouping. */ } return action(); }
    }
}
