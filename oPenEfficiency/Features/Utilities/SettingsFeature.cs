using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnSettings",
        Name = "Settings",
        Tooltip = "Settings",
        Color = "#94A3B8",
        Description = "Configure your oPenEfficiency experience, including sidebar layout, spacing increments, and keyboard shortcuts.",
        DetailedHelpText = "Configure your oPenEfficiency experience, including sidebar layout, spacing increments, and keyboard shortcuts.",
        Keywords = "settings, configuration, preferences, options, customize",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class SettingsFeature
    {
        // Execution handled by sidebar (OpenSettings) to wire up the SettingsApplied -> RebuildUI callback.
        public static bool Execute(PowerPointManager manager) => false;
    }
}
