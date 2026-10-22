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
        Tooltip = "Match selected shapes to the anchor's width and height",
        IconData = "M3,3H9V9H3V3M11,11H21V21H11V11M3,11H9V17H3V11M11,3H17V9H11V3Z",
        Color = "#10B981",
        Description = "Forces all selected objects to match the exact Width and Height of the anchor object.",
        DetailedHelpText = "### Match Size\n\nResizes every non-anchor shape to the exact width and height of the first or last selected anchor configured in Settings. Right-click any Match Size, Match Width, or Match Height button to choose width, height, or both dimensions directly. Existing aspect-ratio lock states are preserved, and the operation can be undone as one PowerPoint step.",
        MinSelection = 2,
        RequiredType = OfficeSelectionType.Shapes)]
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
        /// Full method with parameters for explicit command variants.
        /// </summary>
        public static bool Execute(PowerPointManager manager, bool width = true, bool height = true)
        {
            return manager != null && AlignmentHelpers.MatchSize(manager, width, height);
        }
    }
}
