using System;
using System.Windows.Controls.Primitives;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Utilities
{
    /// <summary>
    /// Snap to Grid feature - toggles snap-to-grid functionality.
    /// When enabled, dragging shapes will snap to grid lines.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnSnapToGrid",
        Name = "Snap to Grid",
        Tooltip = "Snap to grid",
        IconData = "M3,3H21V21H3V3M7,7V9H9V7H7M11,7V9H13V7H11M15,7V9H17V7H15M19,7V9H21V7H19M7,11V13H9V11H7M11,11V13H13V11H11M15,11V13H17V11H15M19,11V13H21V11H19M7,15V17H9V15H7M11,15V17H13V15H11M15,15V17H17V15H15M19,15V17H21V15H19Z",
        Color = "#F43F5E",
        Description = "Toggles snap-to-grid functionality for precise shape positioning.",
        DetailedHelpText = "### Snap to Grid\nEnables session-level grid snapping for one selected shape while it is dragged in PowerPoint. The shape's top-left corner is placed on the nearest 50-point grid intersection and temporary guides are removed after the drag. Enabling this mode disables Snap to Objects to prevent conflicting mouse hooks. Right-click to enable Grid or Object snapping explicitly, or disable both snapping modes.",
        MinSelection = 0,
        IsToggle = true)]
    public static class SnapToGridFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - toggles snap-to-grid functionality.
        /// Note: This feature requires toggle state management, handled by the command adapter.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return SetEnabled(!IsEnabled, null);
        }

        /// <summary>
        /// Toggles the snap-to-grid functionality.
        /// </summary>
        public static void Execute(PowerPointManager manager, ToggleButton toggleBtn = null)
        {
            try
            {
                SetEnabled(!IsEnabled, toggleBtn);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SnapToGridFeature.Execute");
            }
        }

        public static bool IsEnabled => Globals.ThisAddIn?.SnapGridManager?.IsEnabled == true;

        public static bool SetEnabled(bool enabled, ToggleButton toggleBtn = null)
        {
            var host = Globals.ThisAddIn;
            var gridManager = host?.SnapGridManager;
            if (gridManager == null)
            {
                ExceptionLogger.Log("SnapGridManager is not initialized", "SnapToGridFeature.SetEnabled");
                return false;
            }
            if (enabled && host.SnapManager != null) host.SnapManager.IsEnabled = false;
            gridManager.IsEnabled = enabled;
            if (toggleBtn != null) toggleBtn.IsChecked = enabled;
            return true;
        }
    }
}
