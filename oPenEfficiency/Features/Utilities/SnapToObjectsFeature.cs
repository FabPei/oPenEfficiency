using System;
using System.Windows.Controls.Primitives;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Utilities
{
    /// <summary>
    /// Snap to Objects feature - toggles snap-to-objects functionality.
    /// When enabled, dragging shapes will snap to edges and centers of nearby objects.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnSnapToObjects",
        Name = "Snap to Objects",
        Tooltip = "Snap to objects",
        IconData = "M3,3H21V21H3V3M7,7V9H9V7H7M11,7V9H13V7H11M15,7V9H17V7H15M19,7V9H21V7H19M7,11V13H9V11H7M11,11V13H13V11H11M15,11V13H17V11H15M19,11V13H21V11H19M7,15V17H9V15H7M11,15V17H13V15H11M15,15V17H17V15H15M19,15V17H21V15H19Z",
        Color = "#F43F5E",
        Description = "Toggles snap-to-objects functionality for precise shape alignment.",
        DetailedHelpText = "### Snap to Objects\nEnables session-level smart snapping for one selected shape while it is dragged in PowerPoint. Visible shapes on the current slide provide edge and center targets; temporary guides disappear after the drag. Enabling this mode disables Snap to Grid to prevent conflicting mouse hooks. Right-click to enable Object or Grid snapping explicitly, or disable both snapping modes.",
        MinSelection = 0,
        IsToggle = true)]
    public static class SnapToObjectsFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - toggles snap-to-objects functionality.
        /// Note: This feature requires toggle state management, handled by the command adapter.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return SetEnabled(!IsEnabled, null);
        }

        /// <summary>
        /// Toggles the snap-to-objects functionality.
        /// </summary>
        public static void Execute(PowerPointManager manager, ToggleButton toggleBtn = null)
        {
            try
            {
                SetEnabled(!IsEnabled, toggleBtn);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SnapToObjectsFeature.Execute");
            }
        }

        public static bool IsEnabled => Globals.ThisAddIn?.SnapManager?.IsEnabled == true;

        public static bool SetEnabled(bool enabled, ToggleButton toggleBtn = null)
        {
            var host = Globals.ThisAddIn;
            var objectManager = host?.SnapManager;
            if (objectManager == null)
            {
                ExceptionLogger.Log("SnapManager is not initialized", "SnapToObjectsFeature.SetEnabled");
                return false;
            }
            if (enabled && host.SnapGridManager != null) host.SnapGridManager.IsEnabled = false;
            objectManager.IsEnabled = enabled;
            if (toggleBtn != null) toggleBtn.IsChecked = enabled;
            return true;
        }
    }
}
