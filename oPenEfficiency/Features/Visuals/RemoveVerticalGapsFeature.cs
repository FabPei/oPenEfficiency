using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Remove vertical gaps between selected shapes, keeping the master shape in place.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnRemoveVerticalGaps",
        Name = "Remove Vertical Gaps",
        Tooltip = "Remove vertical gaps",
        IconData = "M5,3V5H19V3H5M5,19V21H19V19H5M12,6V18",
        Color = "#10B981",
        Description = "Removes vertical gaps between selected shapes.",
        DetailedHelpText = "### Remove Vertical Gaps\nCollapses all vertical spacing between selected shapes to zero, sliding them together with no gap on the vertical axis.",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class RemoveVerticalGapsFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return RemoveGapsFeatureHelper.Execute(manager, horizontal: false);
        }

        public static bool ExecuteCentered(PowerPointManager manager)
        {
            return RemoveGapsFeatureHelper.Execute(manager, horizontal: false, packToCenter: true);
        }
    }
}
