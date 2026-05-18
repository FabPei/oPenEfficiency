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
        DetailedHelpText = "### Snap to Objects\nToggles smart snapping to nearby shape edges and centers, displaying alignment guides when shapes are dragged near other shapes.",
        MinSelection = 0,
        IsToggle = true)]
    public static class SnapToObjectsFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - toggles snap-to-objects functionality.
        /// Note: This feature requires ToggleButton state management, handled by manual switch.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            // Toggle feature - requires ToggleButton parameter for state sync
            // Handled by manual switch in MainSidebar.xaml.cs
            return false;
        }

        /// <summary>
        /// Toggles the snap-to-objects functionality.
        /// </summary>
        public static void Execute(PowerPointManager manager, ToggleButton toggleBtn = null)
        {
            try
            {
                var snapManager = Globals.ThisAddIn.SnapManager;
                if (snapManager == null)
                {
                    ExceptionLogger.Log("SnapManager is not initialized", "SnapToObjectsFeature.Execute");
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
                ExceptionLogger.Log(ex, "SnapToObjectsFeature.Execute");
            }
        }
    }
}
