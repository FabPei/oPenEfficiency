using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Shapes
{
    [FeatureMetadata(
        Id = "BtnMatchAngles",
        Name = "Apply Anchor Shape Adjustments",
        SearchAliases = new[] { "Match Shape Adjustments", "Match Shape Geometry" },
        Tooltip = "Match yellow-diamond adjustments to the configured anchor",
        IconData = "M2,2H8V4H4V8H2V2M22,2V8H20V4H16V2H22M2,22V16H4V20H8V22H2M22,22H16V20H20V16H22V22M15,10V14L18,12L15,10M9,10L6,12L9,14V10Z",
        Color = "#8B5CF6", // Distinct Purple 
        Description = "Matches adjustable shape geometry to the configured anchor without changing size or position.",
        DetailedHelpText = "### Apply Anchor Shape Adjustments\n\nThis geometry-focused entry uses the same safe adjustment transfer as Copy Shape Adjustments. Select compatible PowerPoint shapes and choose the first or last selected anchor in Settings. Right-click to use the first or last selected shape explicitly without changing Settings. Arrowheads, corner radii, callout tails, and other yellow-diamond values are copied only to shapes with matching geometry; incompatible shapes are skipped, and the complete operation supports Undo and rollback.",
        MinSelection = 2,
        RequiredType = OfficeSelectionType.Shapes,
        UserVisible = false,
        SupportsUserShortcut = false)]
    public static class MatchAnglesFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return ShapeAdjustmentTransfer.Execute(manager, "MatchAnglesFeature.Execute");
        }

        public static bool ExecuteUsingFirstSelected(PowerPointManager manager)
        {
            return ShapeAdjustmentTransfer.Execute(manager, "MatchAnglesFeature.ExecuteUsingFirstSelected", false);
        }

        public static bool ExecuteUsingLastSelected(PowerPointManager manager)
        {
            return ShapeAdjustmentTransfer.Execute(manager, "MatchAnglesFeature.ExecuteUsingLastSelected", true);
        }
    }
}
