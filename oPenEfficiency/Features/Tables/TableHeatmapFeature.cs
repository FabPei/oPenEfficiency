using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Tables
{
    [FeatureMetadata(
        Id = "BtnTableHeatmap",
        Name = "Table Heatmap",
        Tooltip = "Format table with conditional colors",
        IconData = "M3,3H21V21H3V3M5,5V9H9V5H5M10,5V9H14V5H10M15,5V9H19V5H15M5,10V14H9V10H5M10,10V14H14V10H10M15,10V14H19V10H15M5,15V19H9V15H5M10,15V19H14V15H10M15,15V19H19V15H15Z",
        Color = "#F59E0B",
        Description = "Creates a heatmap based on the values. Opens option dialog.",
        DetailedHelpText = "### Table Heatmap\nFormats the selected table as a conditional color heatmap based on numeric cell values.\n**Options:**\n- Choose from 12 Excel-style color scale presets (e.g., Green-Yellow-Red, Blue-White-Red) or define a custom gradient.\n- Apply as a smooth gradient or as icon sets (arrows, circles, flags).\n- Optionally include or exclude the header row from the analysis range.",
        MinSelection = 0,
        RequiresTable = true)]
    public static class TableHeatmapFeature
    {
        public static bool Execute(PowerPointManager manager)
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
                var dialog = new UI.Dialogs.TableHeatmapDialog(manager, tableShape, table);
                dialog.ShowDialog();
                
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableHeatmapFeature.Execute");
                return false;
            }
        }
    }
}
