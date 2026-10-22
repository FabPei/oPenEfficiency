using System;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.UI;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnThemeColor",
        Name = "Theme Color Palette (RGB)",
        SearchAliases = new[] { "Absolute Theme Colors" },
        Tooltip = "Apply current theme colors as fixed RGB values",
        IconData = "M17.5,12A1.5,1.5 0 0,1 16,10.5A1.5,1.5 0 0,1 17.5,9A1.5,1.5 0 0,1 19,10.5A1.5,1.5 0 0,1 17.5,12M14.5,8A1.5,1.5 0 0,1 13,6.5A1.5,1.5 0 0,1 14.5,5A1.5,1.5 0 0,1 16,6.5A1.5,1.5 0 0,1 14.5,8M9.5,8A1.5,1.5 0 0,1 8,6.5A1.5,1.5 0 0,1 9.5,5A1.5,1.5 0 0,1 11,6.5A1.5,1.5 0 0,1 9.5,8M6.5,12A1.5,1.5 0 0,1 5,10.5A1.5,1.5 0 0,1 6.5,9A1.5,1.5 0 0,1 8,10.5A1.5,1.5 0 0,1 6.5,12M12,3A9,9 0 0,0 3,12A9,9 0 0,0 12,21A1.5,1.5 0 0,0 13.5,19.5C13.5,19.11 13.35,18.76 13.11,18.5C12.88,18.23 12.73,17.88 12.73,17.5A1.5,1.5 0 0,1 14.23,16H16A5,5 0 0,0 21,11C21,6.58 16.97,3 12,3Z",
        Color = "#06B6D4",
        Description = "Opens a palette that resolves the active presentation theme to fixed RGB fill, line, and text colors.",
        DetailedHelpText = "### Theme Color Palette (RGB)\n\nOpen the palette, then use the Fill, Border, or Text row to apply the active presentation's theme colors and tint variants as absolute RGB values. Right-click to open the floating palette directly at the Fill, Line, or Text colors, or show and hide the docked palette. The palette refreshes when the active presentation changes, so applied colors do not drift after a later theme change. Custom swatches are session-only fill colors.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class ThemeColorPickerFeature
    {
        private static ThemeColorWindow _currentWindow;

        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;

            try
            {
                if (_currentWindow != null && _currentWindow.IsVisible)
                {
                    _currentWindow.Close();
                    _currentWindow = null;
                    return true;
                }

                return OpenWindow(manager, null);
            }
            catch (Exception ex)
            {
                var failedWindow = _currentWindow;
                _currentWindow = null;
                ExceptionLogger.Log(ex, "ThemeColorPickerFeature.Execute");
                try { failedWindow?.Close(); }
                catch (Exception closeEx)
                {
                    ExceptionLogger.Log(closeEx, "ThemeColorPickerFeature.Execute.FailureCleanup");
                }
                throw;
            }
        }

        public static bool OpenForTarget(PowerPointManager manager, PickColorFeature.ApplyTarget target)
        {
            if (manager == null) return false;
            try
            {
                if (_currentWindow != null && _currentWindow.IsVisible)
                {
                    _currentWindow.FocusTarget(target);
                    return true;
                }
                return OpenWindow(manager, target);
            }
            catch (Exception ex)
            {
                var failedWindow = _currentWindow;
                _currentWindow = null;
                ExceptionLogger.Log(ex, "ThemeColorPickerFeature.OpenForTarget");
                try { failedWindow?.Close(); }
                catch (Exception closeEx)
                {
                    ExceptionLogger.Log(closeEx, "ThemeColorPickerFeature.OpenForTarget.FailureCleanup");
                }
                throw;
            }
        }

        private static bool OpenWindow(PowerPointManager manager, PickColorFeature.ApplyTarget? target)
        {
            _currentWindow = new ThemeColorWindow(manager, target);
            var app = manager.GetApplication();
            if (app != null)
                new System.Windows.Interop.WindowInteropHelper(_currentWindow).Owner = new IntPtr(app.HWND);
            _currentWindow.Closed += (s, e) => { _currentWindow = null; };
            _currentWindow.Show();
            return true;
        }
    }
}
