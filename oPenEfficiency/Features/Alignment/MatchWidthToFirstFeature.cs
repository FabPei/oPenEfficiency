using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Match the width of all selected shapes to the width of the first object.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnMatchWidth",
        Name = "Match Width",
        Tooltip = "Match width",
        IconData = "M3,5H5V19H3V5M19,5H21V19H19V5M6,12H18M8,10L6,12L8,14M16,10L18,12L16,14",
        Color = "#10B981",
        Description = "Standardizes the width of all selected objects to match the anchor object.",
        DetailedHelpText = "### Match Width\nResizes all selected shapes to match the exact pixel width of the first shape, preserving height and position.",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class MatchWidthToFirstFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return manager != null && AlignmentHelpers.MatchSize(manager, matchWidth: true, matchHeight: false);
        }
    }
}
