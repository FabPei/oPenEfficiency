using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using System.Windows.Interop;

namespace oPenEfficiency.Features.Tables
{
    [FeatureMetadata(
        Id = "BtnTableHeatmap",
        Name = "Apply Table Heatmap",
        SearchAliases = new[] { "Apply Table Heatmap Formatting", "Table Heatmap" },
        Tooltip = "Color table cells by their numeric values",
        IconData = "M3,3H21V21H3V3M5,5V9H9V5H5M10,5V9H14V5H10M15,5V9H19V5H15M5,10V14H9V10H5M10,10V14H14V10H10M15,10V14H19V10H15M5,15V19H9V15H5M10,15V19H14V15H10M15,15V19H19V15H15Z",
        Color = "#F59E0B",
        Description = "Applies a row-, column-, or table-scoped color scale or icon prefix to numeric table cells.",
        DetailedHelpText = "### Apply Table Heatmap\n\nChoose row, column, or complete-table normalization; one of 12 color presets or custom colors; fill, font color, or an icon prefix; and whether to exclude the first row or ignore blank cells. A live sample preview shows the active scheme, scope and data options before anything is applied. Right-click opens the dialog with row, column, or entire-table scope preselected. Non-empty non-numeric labels are always left unchanged. Icon mode replaces only a known leading heatmap icon and preserves the remaining cell text. If no numeric values are available nothing changes; successful application creates a PowerPoint Undo entry.",
        MinSelection = 0,
        RequiresTable = true)]
    public static class TableHeatmapFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, UI.Dialogs.TableHeatmapScope.Row);
        }

        public static bool Execute(PowerPointManager manager, UI.Dialogs.TableHeatmapScope initialScope)
        {
            if (manager == null) return false;

            try
            {
                if (!TableHelper.GetSelectedTable(manager, out Shape tableShape, out Table table))
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Please select a table to apply the heatmap to.",
                        "Error",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return false;
                }

                // Initialize and show the dialog
                var dialog = new UI.Dialogs.TableHeatmapDialog(manager, tableShape, table, initialScope);
                var app = manager.GetApplication();
                if (app != null) new WindowInteropHelper(dialog).Owner = (IntPtr)app.HWND;
                dialog.ShowDialog();
                
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableHeatmapFeature.Execute");
                throw;
            }
        }

        /// <summary>
        /// Dialog-free programmatic entry (used by the acceptance seam): applies the heatmap with
        /// explicit options to the selected table. Never raises a message box; returns false when
        /// no table is selected, nothing was applied, or some cells failed.
        /// </summary>
        public static bool Execute(PowerPointManager manager, TableHeatmapOptions options)
        {
            if (manager == null || options == null) return false;
            if (!TableHelper.GetSelectedTable(manager, out Shape tableShape, out Table table)) return false;
            var result = TableHeatmapApplier.Apply(manager, table, options);
            return result.Formatted > 0 && result.Failed == 0 && !result.NoDataRows && !result.NoNumericValues;
        }
    }
}
