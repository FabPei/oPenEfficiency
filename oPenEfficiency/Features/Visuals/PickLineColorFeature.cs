using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnPickLineColor",
        Name = "Pick Line Color",
        Tooltip = "Pick an absolute RGB line color from the screen",
        Color = "#3B82F6",
        Description = "Samples an exact screen pixel and applies its absolute RGB value to outlines and line shapes.",
        DetailedHelpText = "### Pick Line Color\n\nSelect one or more shapes or lines, then choose any visible screen pixel. The sampled absolute RGB value is applied to each compatible outline and makes hidden lines visible. Right-click can redirect the picker to fill, line, or text color. Escape cancels without changes; multi-shape changes use one Undo step and roll back together after an error.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class PickLineColorFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return PickColorFeature.Execute(manager, PickColorFeature.ApplyTarget.Line);
        }
    }
}
