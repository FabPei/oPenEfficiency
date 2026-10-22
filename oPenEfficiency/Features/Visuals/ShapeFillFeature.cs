using System;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.UI.Menus;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Visuals
{
    [FeatureMetadata(
        Id = "BtnShapeFill",
        Name = "Shape Fill",
        SearchAliases = new[] { "Fill Color", "No Fill", "Theme Fill" },
        Tooltip = "Apply theme, standard, recent or custom fills",
        IconData = "M19,11L8,20L3,15L12,6M14,4L20,10M3,21H21",
        Color = "#06B6D4",
        Description = "Opens a compact PowerPoint-style palette for theme-aware, standard, recent and custom shape fills.",
        DetailedHelpText = "Choose a theme color and tint, a standard or recent color, No Fill, More Fill Colors, or Eyedropper. The palette applies one fill operation to all eligible selected shapes and opening it never changes the document.",
        Category = FeatureCategory.TextAndFormat,
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class ShapeFillFeature
    {
        private static ShapeFillWindow _window;

        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                if (_window != null && _window.IsVisible)
                {
                    _window.Close();
                    return true;
                }
                _window = new ShapeFillWindow(manager);
                var app = manager.GetApplication();
                if (app != null)
                    new System.Windows.Interop.WindowInteropHelper(_window).Owner = new IntPtr(app.HWND);
                _window.Closed += (sender, args) => _window = null;
                _window.Show();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ShapeFillFeature.Execute");
                try { _window?.Close(); }
                catch (Exception closeEx) { ExceptionLogger.Log(closeEx, "ShapeFillFeature.Execute.FailureCleanup"); }
                _window = null;
                return false;
            }
        }
    }
}
