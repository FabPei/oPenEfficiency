using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnSizePanel",
        Name = "Size Panel",
        Tooltip = "Shape Size Panel (Width × Height)",
        Color = "#6366F1",
        Description = "Displays and allows editing of exact Width and Height for selected shapes. Supports centimeter and point inputs.",
        DetailedHelpText = "Displays and allows editing of exact Width and Height for selected shapes. Supports centimeter and point inputs.",
        MinSelection = 1,
        IsToggle = true,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class SizePanelFeature
    {
        // Rendered as a full-width widget in the sidebar, not a button.
        public static bool Execute(PowerPointManager manager) => false;
    }
}
