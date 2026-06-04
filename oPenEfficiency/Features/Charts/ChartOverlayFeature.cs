using System;
using System.Reflection;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.UI.Dialogs;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Charts
{
    [FeatureMetadata(
        Id = "BtnChartOverlay",
        Name = "Chart Overlays",
        Tooltip = "Add analytical overlays like CAGR and Delta to native charts",
        IconData = "M16,11.78L20.24,4.45L21.97,5.45L16.74,14.5L10.23,10.75L5.46,19H2.26L8.11,8.38L14.62,12.13L16,11.78M22,12V22H2V10H4V20H20V12H22Z",
        Color = "#10B981",
        Description = "Calculates mathematical offsets over standard charts",
        DetailedHelpText = "### Analytical Chart Overlays\n\nThis feature allows you to cast mathematical VSTO shapes directly over native PowerPoint charts without distorting the layout. \n\n**Features Included:**\n- **CAGR Brackets & Deltas:** Draws precise, colored growth indicators between two targeted data points.\n- **Reference Lines:** Averages your visible chart data to draw a statistical tracking line across your entire plot area.\n- **Outlier Detection:** Highlights data points extending beyond standard deviation boundaries with red ovals.\n\n*Note: Automatically supports stacked column total aggregations.*",
        Keywords = "cagr bracket, delta calculation, statistical line, outlier detection, chart additions",
        MinSelection = 1)]
    public static class ChartOverlayFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;

            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionShapes)
                {
                    Microsoft.Office.Interop.PowerPoint.ShapeRange selection = app.ActiveWindow.Selection.ShapeRange;
                    if (selection.Count == 1 && selection[1].HasChart == MsoTriState.msoTrue)
                    {
                        var chart = selection[1].Chart;
                        ChartOverlayDialog dialog = new ChartOverlayDialog(manager, chart, selection[1]);
                        dialog.Owner = System.Windows.Application.Current.MainWindow;
                        dialog.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner;
                        dialog.ShowDialog();
                        return true;
                    }
                }

                System.Windows.Forms.MessageBox.Show("Please select a single native PowerPoint Chart to add an overlay.", "No Chart Selected", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ChartOverlayFeature.Execute");
                System.Windows.Forms.MessageBox.Show("An error occurred while launching Chart Overlays.", "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }
    }
}
