using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnFontPanel",
        Name = "Font Panel",
        Tooltip = "Font & Text Formatting Panel",
        IconData = "M12 4v16 M4 7V5a1 1 0 0 1 1-1h14a1 1 0 0 1 1 1v2 M9 20h6",
        Color = "#EAB308",
        Description = "Toggles an inline panel for quickly changing font family, size, and text formatting options.",
        DetailedHelpText = "### Font Panel\nExpands an inline panel directly in the sidebar for comprehensive text formatting. Adjust font family, size, bold, italic, underline, and more — all without leaving the sidebar.",
        Keywords = "text, typography, style, family, inline, format",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone,
        IsToggle = true)]
    public static class FontPanelFeature
    {
        public static bool Execute(PowerPointManager manager) => false;
    }
}
