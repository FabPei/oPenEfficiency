using System;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnPerformanceMode",
        Name = "Performance Mode",
        SearchAliases = new[] { "Pause External Add-ins", "COM Add-ins" },
        Tooltip = "Temporarily pause selected external PowerPoint add-ins",
        IconData = "M13,2L3,14H11L10,22L21,9H13V2Z",
        Color = "#F59E0B",
        Description = "Temporarily pauses only the external COM add-ins explicitly selected by the user and restores only those changed in this session.",
        DetailedHelpText = "Right-click to choose detected external PowerPoint COM add-ins. Nothing is selected by default. Enabling Performance Mode pauses selected active add-ins through their supported Connect property; disabling or shutting down restores only add-ins this session paused.",
        Category = FeatureCategory.SettingsAndHelp,
        RequiredType = OfficeSelectionType.None,
        IsToggle = true,
        SupportsUserShortcut = false)]
    public static class ExternalAddInPerformanceModeFeature
    {
        public static bool IsEnabled(PowerPointManager manager) =>
            manager != null && Services.ExternalAddInPerformanceModeService.For(manager.GetApplication()).IsEnabled;

        public static bool Execute(PowerPointManager manager) => SetEnabled(manager, !IsEnabled(manager));

        public static bool SetEnabled(PowerPointManager manager, bool enabled)
        {
            try
            {
                return manager != null &&
                    Services.ExternalAddInPerformanceModeService.For(manager.GetApplication()).SetEnabled(enabled);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExternalAddInPerformanceModeFeature.SetEnabled");
                return false;
            }
        }
    }
}
