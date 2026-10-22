using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnPickFillColor",
        Name = "Pick Fill Color",
        Tooltip = "Pick an absolute RGB fill color from the screen",
        Color = "#06B6D4",
        Description = "Samples an exact screen pixel and applies its absolute RGB value to the fill foreground of selected shapes.",
        DetailedHelpText = "### Pick Fill Color\n\nSelect one or more shapes, then choose any visible screen pixel. The sampled absolute RGB value is applied to each compatible fill foreground while the existing fill style is preserved. Right-click can redirect the picker to fill, line, or text color. Escape cancels without changes; multi-shape changes use one Undo step and roll back together after an error.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class PickFillColorFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return PickColorFeature.Execute(manager, PickColorFeature.ApplyTarget.Fill);
        }
    }
}
