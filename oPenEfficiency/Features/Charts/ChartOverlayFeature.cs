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
        Name = "Add Analytical Chart Overlays",
        SearchAliases = new[] { "Chart Overlays" },
        Tooltip = "Add analytical overlays like CAGR and Delta to native charts or chart groups",
        IconData = "M16,11.78L20.24,4.45L21.97,5.45L16.74,14.5L10.23,10.75L5.46,19H2.26L8.11,8.38L14.62,12.13L16,11.78M22,12V22H2V10H4V20H20V12H22Z",
        Color = "#10B981",
        Description = "Calculates mathematical offsets over standard charts, including a chart inside a selected group",
        DetailedHelpText = "### Add Analytical Chart Overlays\n\nSelect exactly one native PowerPoint chart or one group containing exactly one chart. Right-click preselects Delta, CAGR, Growth, Average Level Line, or Highlight Outliers; the dialog still lets you review point ranges and styling before insertion. The chart remains grouped, while the generated overlay is added to the slide as an independent grouped annotation. Overlays are mathematical PowerPoint shapes and support stacked-column total aggregation.",
        MinSelection = 1)]
    public static class ChartOverlayFeature
    {
        public enum OverlayPreset
        {
            Default,
            Delta,
            Cagr,
            Growth,
            AverageLevel,
            Outliers
        }

        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, OverlayPreset.Default);
        }

        public static bool Execute(PowerPointManager manager, OverlayPreset preset)
        {
            if (manager == null) return false;

            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow.Selection.Type == PpSelectionType.ppSelectionShapes)
                {
                    Microsoft.Office.Interop.PowerPoint.ShapeRange selection = app.ActiveWindow.Selection.ShapeRange;
                    if (selection.Count == 1)
                    {
                        Microsoft.Office.Interop.PowerPoint.Shape chartShape = FindSingleChart(selection[1]);
                        if (chartShape != null)
                        {
                            var chart = chartShape.Chart;
                            ChartOverlayDialog dialog = new ChartOverlayDialog(manager, chart, chartShape, preset);
                            dialog.Owner = System.Windows.Application.Current.MainWindow;
                            dialog.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner;
                            dialog.ShowDialog();
                            return true;
                        }
                    }
                }

                System.Windows.Forms.MessageBox.Show("Select one native PowerPoint chart or one group containing exactly one chart.", "No Single Chart Selected", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ChartOverlayFeature.Execute");
                System.Windows.Forms.MessageBox.Show("An error occurred while launching Add Analytical Chart Overlays.", "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }

        private static Microsoft.Office.Interop.PowerPoint.Shape FindSingleChart(
            Microsoft.Office.Interop.PowerPoint.Shape selectedShape)
        {
            Microsoft.Office.Interop.PowerPoint.Shape chartShape = null;
            int chartCount = 0;
            CollectCharts(selectedShape, ref chartShape, ref chartCount);
            return chartCount == 1 ? chartShape : null;
        }

        private static void CollectCharts(
            Microsoft.Office.Interop.PowerPoint.Shape shape,
            ref Microsoft.Office.Interop.PowerPoint.Shape chartShape,
            ref int chartCount)
        {
            if (shape == null || chartCount > 1) return;

            if (shape.HasChart == MsoTriState.msoTrue)
            {
                chartShape = shape;
                chartCount++;
                return;
            }

            if (shape.Type != MsoShapeType.msoGroup) return;
            for (int index = 1; index <= shape.GroupItems.Count; index++)
                CollectCharts(shape.GroupItems[index], ref chartShape, ref chartCount);
        }
    }
}
