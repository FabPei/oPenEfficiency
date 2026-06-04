using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Match the size of selected shapes to the first/master object (Width and Height).
    /// </summary>
    [FeatureMetadata(
        Id = "BtnMatchSize",
        Name = "Match Size",
        Tooltip = "Match size",
        IconData = "M3,3H9V9H3V3M11,11H21V21H11V11M3,11H9V17H3V11M11,3H17V9H11V3Z",
        Color = "#10B981",
        Description = "Forces all selected objects to match the exact Width and Height of the anchor object.",
        DetailedHelpText = "### Match Size\nResizes all selected shapes to match the exact width AND height of the first shape simultaneously.",
        Keywords = "same size, copy dimensions, identical size, match width and height",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class MatchSizeFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - matches both width and height (default behavior).
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return manager != null && AlignmentHelpers.MatchSize(manager, true, true);
        }

        /// <summary>
        /// Full method with parameters for manual switch usage.
        /// </summary>
        public static bool Execute(PowerPointManager manager, bool width = true, bool height = true)
        {
            return manager != null && AlignmentHelpers.MatchSize(manager, width, height);
        }
    }
}
