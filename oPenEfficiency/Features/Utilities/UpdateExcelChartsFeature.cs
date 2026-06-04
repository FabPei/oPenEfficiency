using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Utilities
{
    /// <summary>
    /// Feature: Updates all Excel-linked charts in the presentation.
    /// Can update charts in selected slide only or all slides in the presentation.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnUpdateExcelCharts",
        Name = "Update Excel Charts",
        Tooltip = "Update Excel charts",
        IconData = "M3,3H21V21H3V3M5,5V19H19V5H5M7,7H10V10H7V7M11,7H14V10H11V7M15,7H18V10H15V7M7,11H10V14H7V11M11,11H14V14H11V11M15,11H18V14H15V11M7,15H10V18H7V15M11,15H14V18H11V15M15,15H18V18H15V15M19.14,17.88L20.5,19.25L18.31,21.44L16.94,20.08L19.14,17.88Z",
        Color = "#F43F5E",
        Description = "Updates all Excel-linked charts in the selected slide.",
        DetailedHelpText = "### Update Excel Charts\nForces a data refresh on all embedded Excel-linked charts in the presentation, pulling the latest values from their source workbooks.",
        Keywords = "excel, charts, update, refresh, data link, sync",
        MinSelection = 0,
        RequiredType = PpSelectionType.ppSelectionNone)]
    public static class UpdateExcelChartsFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - updates Excel-linked charts in selected slide.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, updateAllSlides: false);
        }

        /// <summary>
        /// Updates Excel-linked charts, optionally across all slides.
        /// </summary>
        /// <param name="manager">PowerPoint manager</param>
        /// <param name="updateAllSlides">If true, updates all slides; otherwise, only selected slide</param>
        public static bool Execute(PowerPointManager manager, bool updateAllSlides)
        {
            if (manager == null) return false;

            try
            {
                var app = manager.GetApplication();
                if (app == null || app.ActiveWindow == null || app.ActivePresentation == null)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "No active presentation found.",
                        "Error",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return false;
                }

                var presentation = app.ActivePresentation;
                var slidesToUpdate = new List<Slide>();

                if (updateAllSlides)
                {
                    // Count charts first to warn user
                    int totalCharts = 0;
                    foreach (Slide slide in presentation.Slides)
                    {
                        totalCharts += CountExcelLinkedCharts(slide);
                    }

                    // Warn user if there are many charts
                    if (totalCharts > 10)
                    {
                        var result = System.Windows.Forms.MessageBox.Show(
                            $"This presentation contains {totalCharts} Excel-linked charts.\n\n" +
                            $"Updating all charts may take a while and will open Excel temporarily.\n\n" +
                            $"Click OK to proceed, or Cancel to update only the selected slide.",
                            "Confirm Update All Charts",
                            System.Windows.Forms.MessageBoxButtons.OKCancel,
                            System.Windows.Forms.MessageBoxIcon.Warning);

                        if (result == System.Windows.Forms.DialogResult.Cancel)
                        {
                            updateAllSlides = false;
                        }
                    }

                    if (updateAllSlides)
                    {
                        foreach (Slide slide in presentation.Slides)
                        {
                            slidesToUpdate.Add(slide);
                        }
                    }
                    else
                    {
                        // Fall back to selected slide
                        var currentSlide = (Slide)app.ActiveWindow.View.Slide;
                        if (currentSlide != null)
                        {
                            slidesToUpdate.Add(currentSlide);
                        }
                    }
                }
                else
                {
                    // Only update selected slide
                    var currentSlide = (Slide)app.ActiveWindow.View.Slide;
                    if (currentSlide != null)
                    {
                        slidesToUpdate.Add(currentSlide);
                    }
                }

                // Update charts in selected slides
                int updatedCount = 0;
                foreach (var slide in slidesToUpdate)
                {
                    updatedCount += UpdateChartsInSlide(slide);
                }

                if (updatedCount == 0)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "No Excel-linked charts found to update.",
                        "Info",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show(
                        $"Successfully updated {updatedCount} Excel-linked chart(s).",
                        "Complete",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "UpdateExcelChartsFeature.Execute");
                System.Windows.Forms.MessageBox.Show(
                    $"Error updating charts: {ex.Message}",
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// Counts the number of Excel-linked charts in a slide.
        /// </summary>
        private static int CountExcelLinkedCharts(Slide slide)
        {
            try
            {
                int count = 0;
                foreach (Shape shape in slide.Shapes)
                {
                    if (IsExcelLinkedChart(shape))
                    {
                        count++;
                    }
                }
                return count;
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// Updates all Excel-linked charts in a slide.
        /// </summary>
        private static int UpdateChartsInSlide(Slide slide)
        {
            try
            {
                int updatedCount = 0;

                foreach (Shape shape in slide.Shapes)
                {
                    if (IsExcelLinkedChart(shape))
                    {
                        UpdateChart(shape);
                        updatedCount++;
                    }
                }

                return updatedCount;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "UpdateChartsInSlide");
                return 0;
            }
        }

        /// <summary>
        /// Checks if a shape is an Excel-linked chart.
        /// </summary>
        private static bool IsExcelLinkedChart(Shape shape)
        {
            try
            {
                // Check if shape has a chart
                if (shape.HasChart == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    // Check if chart has linked data
                    var chart = shape.Chart;
                    if (chart != null)
                    {
                        // Access ChartData to check if it's linked to Excel
                        var chartData = chart.ChartData;
                        if (chartData != null)
                        {
                            // Try to get the workbook - if it exists, it's Excel-linked
                            var workbook = chartData.Workbook as Microsoft.Office.Interop.Excel.Workbook;
                            if (workbook != null)
                            {
                                return true;
                            }
                        }
                    }
                }
                return false;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Updates a single Excel-linked chart.
        /// </summary>
        private static void UpdateChart(Shape shape)
        {
            try
            {
                var chart = shape.Chart;
                if (chart == null) return;

                // Access the chart data to trigger refresh
                var chartData = chart.ChartData;
                if (chartData == null) return;

                // Get workbook reference (opens Excel if not already open)
                var workbook = chartData.Workbook as Microsoft.Office.Interop.Excel.Workbook;
                if (workbook == null) return;

                // Refresh the chart by re-querying the data
                chart.Refresh();

                // Close the Excel workbook if we opened it (cleanup)
                // Note: We don't save, just close to release the COM object
                try
                {
                    workbook.Close(false);
                }
                catch
                {
                    // Ignore errors on close - Excel might be used by other charts
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "UpdateChart");
            }
        }
    }
}
