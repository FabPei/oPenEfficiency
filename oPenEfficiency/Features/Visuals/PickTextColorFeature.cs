using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnPickTextColor",
        Name = "Pick Text Color",
        Tooltip = "Pick an absolute RGB text color from the screen",
        Color = "#F59E0B",
        Description = "Samples an exact screen pixel and applies its absolute RGB value to all text in compatible selected shapes.",
        DetailedHelpText = "### Pick Text Color\n\nSelect one or more text-capable shapes, then choose any visible screen pixel. The sampled absolute RGB value replaces the text color across each compatible selected shape; non-text shapes are skipped. Right-click can redirect the picker to fill, line, or text color. Escape cancels without changes, and partial failures restore captured rich-text formatting where possible.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class PickTextColorFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return PickColorFeature.Execute(manager, PickColorFeature.ApplyTarget.Text);
        }
    }
}
