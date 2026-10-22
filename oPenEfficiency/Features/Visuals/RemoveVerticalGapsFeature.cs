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
        DetailedHelpText = "### Remove Vertical Gaps\nSelect at least two shapes. The command removes the vertical gaps while keeping the first selected shape in place and preserving the top-to-bottom order. Right-click the sidebar button and choose **Remove Vertical Gaps & Center** to center the packed group within its original vertical bounds instead.",
        MinSelection = 2,
        RequiredType = OfficeSelectionType.Shapes)]
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
