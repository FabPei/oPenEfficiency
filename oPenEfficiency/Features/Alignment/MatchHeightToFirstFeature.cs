using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Match the height of all selected shapes to the height of the first object.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnMatchHeight",
        Name = "Match Height",
        Tooltip = "Match height",
        IconData = "M5,3V5H19V3H5M5,19V21H19V19H5M12,6V18M10,8L12,6L14,8M10,16L12,18L14,16",
        Color = "#10B981",
        Description = "Standardizes the height of all selected objects to match the anchor object.",
        DetailedHelpText = "### Match Height\nResizes all selected shapes to match the exact pixel height of the first shape, preserving width and position.",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class MatchHeightToFirstFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return manager != null && AlignmentHelpers.MatchSize(manager, matchWidth: false, matchHeight: true);
        }
    }
}
