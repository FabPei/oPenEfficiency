using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.Features.Utilities
{
    /// <summary>
    /// Feature: Updates all Excel-linked charts in the presentation.
    /// Can update charts in selected slide only or all slides in the presentation.
    /// </summary>
    public static class UpdateExcelChartsFeature
    {
        /// <summary>
        /// Shortcut/manual entry point: updates Excel-linked charts on the current slide.
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
                            $"Updating all charts may take a while.\n\n" +
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
            => Services.ExcelChartUpdateService.CountLinkedCharts(slide);

        /// <summary>
        /// Updates all Excel-linked charts in a slide.
        /// </summary>
        private static int UpdateChartsInSlide(Slide slide)
            => Services.ExcelChartUpdateService.UpdateChartsInSlide(slide);


    }
}
