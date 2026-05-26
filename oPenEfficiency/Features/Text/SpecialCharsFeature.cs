using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnSpecialChars",
        Name = "Special Characters",
        Tooltip = "Special Characters",
        Color = "#FDE047",
        Description = "Opens a specialized picker for inserting symbols, currencies, and mathematical operators not easily found on the keyboard.",
        DetailedHelpText = "Opens a specialized picker for inserting symbols and special characters.",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class SpecialCharsFeature
    {
        // Execution is handled by the sidebar (OpenSpecialCharactersWindow) for window positioning.
        public static bool Execute(PowerPointManager manager) => false;
    }
}
