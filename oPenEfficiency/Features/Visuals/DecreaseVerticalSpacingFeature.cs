using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnDecreaseVerticalSpacing",
        Name = "Decrease Vertical Spacing",
        Tooltip = "Decrease vertical spacing between shapes",
        IconData = "M12 22v-6 M12 8V2 M4 12H2 M10 12H8 M16 12h-2 M22 12h-2 M15 19l-3-3-3 3 M15 5l-3 3-3-3",
        Color = "#10B981",
        Description = "Decreases the vertical gap between selected shapes by the configured spacing increment.",
        DetailedHelpText = "### Decrease Vertical Spacing\nMoves selected shapes closer together vertically by the spacing increment (configurable in Settings > Spacing Increment). Select 2 or more shapes and click to compress their spacing.",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class DecreaseVerticalSpacingFeature
    {
        public static bool Execute(PowerPointManager manager) => false;
    }
}
