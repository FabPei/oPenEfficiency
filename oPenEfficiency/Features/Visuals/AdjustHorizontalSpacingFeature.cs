using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Adjust the horizontal spacing between selected shapes.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnAdjustSpacing",
        Name = "Adjust Spacing",
        Tooltip = "Adjust spacing",
        IconData = "M21,11H13V3H11V11H3V13H11V21H13V13H21V11Z",
        Color = "#10B981",
        Description = "Adjusts horizontal spacing between selected shapes (default increment).",
        DetailedHelpText = "### Adjust Spacing (Horizontal)\nIncrements or decrements the horizontal gap between all selected shapes by a fixed amount per button press.",
        Keywords = "spacing, horizontal, distribute, gap, distance",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class AdjustHorizontalSpacingFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - adjusts horizontal spacing with default increment.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, 5f);
        }

        public static bool Execute(PowerPointManager manager, float increment = 5f)
        {
            return AdjustSpacingFeatureHelper.Execute(manager, increment, horizontal: true);
        }
    }
}
