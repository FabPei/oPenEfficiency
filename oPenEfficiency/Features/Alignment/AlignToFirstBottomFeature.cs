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
        Tooltip = "Align bottom edges or snap to the bottom guide",
        IconData = "M22,20V22H2V20H22M10,2V18H14V2H10Z",
        Color = "#10B981",
        Description = "Multi-select: aligns bottom edges to the anchor object. Single select: snaps bottom edge to the nearest horizontal guide line (or slide edge). Right-click for 'Move Until Object'.",
        DetailedHelpText = "### Align Bottom\n\nWith multiple shapes, aligns their bottom edges to the first or last selected anchor configured in Settings. With one shape, snaps its bottom edge to the nearest lower horizontal guide or to the slide bottom. Right-click for explicit selection, slide, first-anchor, last-anchor, and collision modes. Each action creates a PowerPoint Undo entry.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class AlignToFirstBottomFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return AlignmentHelpers.ExecuteAlignment(manager, "bottom");
        }

        public static bool ExecuteWithinSelection(PowerPointManager manager) => ExecuteExplicit(manager, () => AlignmentHelpers.AlignShapes(manager, "bottom"));
        public static bool ExecuteToSlide(PowerPointManager manager) => ExecuteExplicit(manager, () => AlignmentHelpers.AlignShapesToSlide(manager, "bottom"));
        public static bool ExecuteToFirstSelected(PowerPointManager manager) => ExecuteExplicit(manager, () => AlignmentHelpers.AlignToSelectedAnchor(manager, "bottom", false));
        public static bool ExecuteToLastSelected(PowerPointManager manager) => ExecuteExplicit(manager, () => AlignmentHelpers.AlignToSelectedAnchor(manager, "bottom", true));
        private static bool ExecuteExplicit(PowerPointManager manager, System.Func<bool> action) { try { manager?.GetApplication()?.StartNewUndoEntry(); } catch { /* Older PowerPoint hosts may not expose undo grouping. */ } return action(); }
    }
}
