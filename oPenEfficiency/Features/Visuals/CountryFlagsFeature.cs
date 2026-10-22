using System;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.UI.Dialogs;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Visuals
{
    [FeatureMetadata(
        Id = "BtnCountryFlags",
        Name = "Country Flags",
        SearchAliases = new[] { "Flag Library", "ISO Flags", "Insert Flag" },
        Tooltip = "Search and insert an offline ISO country flag as SVG",
        IconData = "M5,21V4M5,5H17L15,9L17,13H5",
        Color = "#3B82F6",
        Description = "Searches the bundled ISO 3166-1 flag catalog and inserts the selected flag as vector SVG content.",
        DetailedHelpText = "Search by country name, alpha-2, alpha-3 or a common alias. Recent flags appear first. Enter or double-click inserts the bundled SVG at its native aspect ratio in one PowerPoint Undo step; normal use is fully offline.",
        Category = FeatureCategory.Libraries,
        RequiredType = OfficeSelectionType.None)]
    public static class CountryFlagsFeature
    {
        private static CountryFlagPickerWindow _window;

        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                if (_window != null && _window.IsVisible) { _window.Activate(); return true; }
                _window = new CountryFlagPickerWindow(manager);
                var app = manager.GetApplication();
                if (app != null) new System.Windows.Interop.WindowInteropHelper(_window).Owner = new IntPtr(app.HWND);
                _window.Closed += (sender, args) => _window = null;
                _window.Show();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "CountryFlagsFeature.Execute");
                return false;
            }
        }
    }
}
