using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Remove horizontal gaps between selected shapes, keeping the master shape in place.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnRemoveHorizontalGaps",
        Name = "Remove Horizontal Gaps",
        Tooltip = "Remove horizontal gaps",
        IconData = "M3,5H5V19H3V5M19,5H21V19H19V5M6,12H18",
        Color = "#10B981",
        Description = "Removes horizontal gaps between selected shapes.",
        DetailedHelpText = "### Remove Horizontal Gaps\nCollapses all horizontal spacing between selected shapes to zero, sliding them together with no gap on the horizontal axis.",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class RemoveHorizontalGapsFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return RemoveGapsFeatureHelper.Execute(manager, horizontal: true);
        }

        public static bool ExecuteCentered(PowerPointManager manager)
        {
            return RemoveGapsFeatureHelper.Execute(manager, horizontal: true, packToCenter: true);
        }
    }
}
