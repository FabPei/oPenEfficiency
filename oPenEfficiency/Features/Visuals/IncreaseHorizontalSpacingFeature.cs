using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnIncreaseHorizontalSpacing",
        Name = "Increase Horizontal Spacing",
        Tooltip = "Increase horizontal spacing between shapes",
        IconData = "M16 12h6 M8 12H2 M12 2v2 M12 8v2 M12 14v2 M12 20v2 M19 15l3-3-3-3 M5 9l-3 3 3 3",
        Color = "#10B981",
        Description = "Increases the horizontal gap between selected shapes by the configured spacing increment.",
        DetailedHelpText = "### Increase Horizontal Spacing\nMoves selected shapes apart horizontally by the spacing increment (configurable in Settings > Spacing Increment). Select 2 or more shapes and click to spread them out.",
        Keywords = "spacing, horizontal, apart, gap, increase",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class IncreaseHorizontalSpacingFeature
    {
        public static bool Execute(PowerPointManager manager) => false;
    }
}
