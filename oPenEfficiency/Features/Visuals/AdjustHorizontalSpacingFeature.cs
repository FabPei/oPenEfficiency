using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Adjust the horizontal spacing between selected shapes.
    /// </summary>
    public static class AdjustHorizontalSpacingFeature
    {
        /// <summary>
        /// Adjusts horizontal spacing with the supplied increment. BtnAdjustSpacing is
        /// registered manually because its primary click opens a direction menu.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, 5f);
        }

        public static bool Execute(PowerPointManager manager, float increment = 5f)
        {
            return AdjustSpacingFeatureHelper.Execute(manager, increment, horizontal: true);
        }

        public static bool ExecuteFromCenter(PowerPointManager manager, float increment = -5f)
        {
            return AdjustSpacingFeatureHelper.Execute(
                manager,
                increment,
                horizontal: true,
                preserveCenter: true);
        }
    }
}
