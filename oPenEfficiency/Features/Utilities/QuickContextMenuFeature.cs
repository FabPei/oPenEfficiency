using System;
using oPenEfficiency.Models;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnQuickContextMenu",
        Name = "Quick Context Menu",
        SearchAliases = new[] { "Quick Menu", "Right-click menu", "Context actions" },
        Tooltip = "Toggle a compact openEfficiency menu beside PowerPoint's native menu",
        IconData = "M3,4H21V8H3V4M3,10H21V14H3V10M3,16H15V20H3V16Z",
        Color = "#6366F1",
        Description = "Shows an opt-in, configurable set of openEfficiency commands beside PowerPoint's native shape context menu.",
        DetailedHelpText = "Click to enable or disable the Quick Context Menu. Right-click this Sidebar command to choose and order up to 12 commands. PowerPoint's native context menu remains available.",
        Category = FeatureCategory.SettingsAndHelp,
        RequiredType = OfficeSelectionType.None,
        IsToggle = true,
        DefaultPlacement = FeatureDefaultPlacement.ProfessionalDefault,
        SupportsUserShortcut = false)]
    public static class QuickContextMenuFeature
    {
        public static bool IsEnabled =>
            JsonService.Load<SidebarConfig>("sidebar_config.json")?.QuickContextMenuEnabled == true;

        public static bool Execute(PowerPointManager manager)
        {
            try
            {
                var config = JsonService.Load<SidebarConfig>("sidebar_config.json") ?? new SidebarConfig();
                config.QuickContextMenuEnabled = !config.QuickContextMenuEnabled;
                if (config.QuickContextMenuFeatureIds == null || config.QuickContextMenuFeatureIds.Count == 0)
                    config.QuickContextMenuFeatureIds = new System.Collections.Generic.List<string>(QuickContextMenuSettings.DefaultFeatureIds);
                return JsonService.TrySave(config, "sidebar_config.json");
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "QuickContextMenuFeature.Execute");
                return false;
            }
        }

        public static bool SetEnabled(bool enabled)
        {
            try
            {
                var config = JsonService.Load<SidebarConfig>("sidebar_config.json") ?? new SidebarConfig();
                config.QuickContextMenuEnabled = enabled;
                if (config.QuickContextMenuFeatureIds == null || config.QuickContextMenuFeatureIds.Count == 0)
                    config.QuickContextMenuFeatureIds = new System.Collections.Generic.List<string>(QuickContextMenuSettings.DefaultFeatureIds);
                return JsonService.TrySave(config, "sidebar_config.json");
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "QuickContextMenuFeature.SetEnabled");
                return false;
            }
        }
    }
}
