using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnSearchBar",
        Name = "Search Bar",
        Tooltip = "Search Command Palette",
        Color = "#10B981",
        Description = "Instantly find and execute any oPenEfficiency command or feature by name using a keyboard-first search bar.",
        DetailedHelpText = "Instantly find and execute any command or feature by name.",
        Keywords = "search, command palette, find, execute, quick access",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class SearchBarFeature
    {
        // Rendered as a full-width widget in the sidebar, not a button.
        public static bool Execute(PowerPointManager manager) => false;
    }
}
