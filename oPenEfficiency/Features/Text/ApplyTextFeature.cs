using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnApplyTextTool",
        Name = "Apply Text",
        Tooltip = "Apply text to all selected shapes",
        IconData = "M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2 M8 2h8v4H8z M9 12v-1h6v1 M11 17h2 M12 11v6",
        Color = "#EAB308",
        Description = "Opens a floating input window to type text and apply it to all selected shapes at once.",
        DetailedHelpText = "### Apply Text\nOpens a compact floating window. Type your text and click Apply (or press Enter) to replace the text in every selected shape simultaneously.\n\n**Options:**\n* **Apply**: Replaces the text in all selected shapes.\n* **Apply to master**: Sets the text on the first (anchor) shape only.",
        Keywords = "bulk text, edit text, mass text, replace text, write multiple",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes,
        IsToggle = true)]
    public static class ApplyTextFeature
    {
        public static bool Execute(PowerPointManager manager) => false;
    }
}
