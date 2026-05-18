using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Adjust the vertical spacing between selected shapes.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnAdjustSpacing",
        Name = "Adjust Spacing",
        Tooltip = "Adjust spacing",
        IconData = "M21,11H13V3H11V11H3V13H11V21H13V13H21V11Z",
        Color = "#10B981",
        Description = "Adjusts vertical spacing between selected shapes (default increment).",
        DetailedHelpText = "### Adjust Spacing (Vertical)\nIncrements or decrements the vertical gap between all selected shapes by a fixed amount per button press.",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class AdjustVerticalSpacingFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - adjusts vertical spacing with default increment.
        /// Note: This shares the same button ID as AdjustHorizontalSpacingFeature - only one can be in the sidebar.
        /// The helper class handles both directions via context menu.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, 5f);
        }

        public static bool Execute(PowerPointManager manager, float increment = 5f)
        {
            return AdjustSpacingFeatureHelper.Execute(manager, increment, horizontal: false);
        }
    }
}
