using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnDecreaseHorizontalSpacing",
        Name = "Decrease Horizontal Spacing",
        Tooltip = "Decrease horizontal spacing between shapes",
        IconData = "M2 12h6 M22 12h-6 M12 2v2 M12 8v2 M12 14v2 M12 20v2 M19 9l-3 3 3 3 M5 15l3-3-3-3",
        Color = "#10B981",
        Description = "Decreases the horizontal gap between selected shapes by the configured spacing increment.",
        DetailedHelpText = "### Decrease Horizontal Spacing\nMoves selected shapes closer together horizontally by the spacing increment (configurable in Settings > Spacing Increment). Select 2 or more shapes and click to compress their spacing.",
        MinSelection = 2,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class DecreaseHorizontalSpacingFeature
    {
        public static bool Execute(PowerPointManager manager) => false;
    }
}
