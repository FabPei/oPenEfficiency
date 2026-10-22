using System;
using oPenEfficiency.Excel.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Excel.Features.View
{
    /// <summary>
    /// Toggles Excel's formula view: shows the underlying formula in every cell
    /// instead of the calculated value. Clicking again restores the values.
    /// Uses the native Window.DisplayFormulas property, so it is fully reversible
    /// and does not modify any cell content.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnShowFormulas",
        Name = "Show Formulas",
        Tooltip = "Show formulas instead of cell values",
        IconData = "",
        Color = "#6366F1",
        Description = "Displays the formula used in each cell instead of the calculated value. Click again to restore the cell values.",
        RequiredType = OfficeSelectionType.None,
        IsToggle = true)]
    public static class ShowFormulasFeature
    {
        /// <summary>Whether the active window currently shows formulas instead of values.</summary>
        public static bool IsShowingFormulas(ExcelManager manager)
        {
            try
            {
                var window = manager?.Application?.ActiveWindow;
                return window != null && window.DisplayFormulas;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>Explicitly sets whether formulas (true) or values (false) are shown.</summary>
        public static void SetShowFormulas(ExcelManager manager, bool showFormulas)
        {
            try
            {
                var window = manager?.Application?.ActiveWindow;
                if (window != null)
                {
                    window.DisplayFormulas = showFormulas;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ShowFormulasFeature.SetShowFormulas");
            }
        }

        /// <summary>Toggles between formula view and value view.</summary>
        public static bool Execute(ExcelManager manager)
        {
            try
            {
                var window = manager?.Application?.ActiveWindow;
                if (window == null)
                {
                    return false;
                }

                window.DisplayFormulas = !window.DisplayFormulas;
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ShowFormulasFeature.Execute");
                return false;
            }
        }
    }
}
