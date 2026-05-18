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
        DetailedHelpText = "### Snap to Grid\nToggles the snap-to-grid behavior. When enabled, shapes are automatically magnetized to the nearest grid point when moved or resized.",
        MinSelection = 0,
        IsToggle = true)]
    public static class SnapToGridFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - toggles snap-to-grid functionality.
        /// Note: This feature requires ToggleButton state management, handled by manual switch.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            // Toggle feature - requires ToggleButton parameter for state sync
            // Handled by manual switch in MainSidebar.xaml.cs
            return false;
        }

        /// <summary>
        /// Toggles the snap-to-grid functionality.
        /// </summary>
        public static void Execute(PowerPointManager manager, ToggleButton toggleBtn = null)
        {
            try
            {
                var snapManager = Globals.ThisAddIn.SnapGridManager;
                if (snapManager == null)
                {
                    ExceptionLogger.Log("SnapGridManager is not initialized", "SnapToGridFeature.Execute");
                    return;
                }

                // Toggle the state
                snapManager.Toggle();

                // Sync the toggle button state with the actual manager state
                if (toggleBtn != null)
                {
                    toggleBtn.IsChecked = snapManager.IsEnabled;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SnapToGridFeature.Execute");
            }
        }
    }
}
