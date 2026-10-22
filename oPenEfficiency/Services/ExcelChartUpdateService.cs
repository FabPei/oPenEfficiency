using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Modal-free update mechanics for Excel-linked charts (QA-S9 wave 3): which shapes count as
    /// Excel-linked charts and how they are refreshed. Extracted from UpdateExcelChartsFeature so
    /// the acceptance seams can drive the production update path without the feature's
    /// success/zero-result MessageBoxes; the feature keeps the reporting on top.
    /// </summary>
    internal static class ExcelChartUpdateService
    {
        /// <summary>
        /// A shape counts as an Excel-linked chart only when it has a chart AND is a linked OLE
        /// object AND carries a non-empty link source. An embedded chart deliberately does not
        /// count - LinkFormat.Update() would not refresh it from its workbook.
        /// </summary>
        public static bool IsLinkedChart(Shape shape)
        {
            try
            {
                return shape != null &&
                       shape.HasChart == Office.MsoTriState.msoTrue &&
                       shape.Type == Office.MsoShapeType.msoLinkedOLEObject &&
                       shape.LinkFormat != null &&
                       !string.IsNullOrWhiteSpace(shape.LinkFormat.SourceFullName);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExcelChartUpdateService.IsLinkedChart");
                return false;
            }
        }

        /// <summary>Counts the Excel-linked charts on a slide.</summary>
        public static int CountLinkedCharts(Slide slide)
        {
            try
            {
                int count = 0;
                foreach (Shape shape in slide.Shapes)
                {
                    if (IsLinkedChart(shape))
                    {
                        count++;
                    }
                }
                return count;
            }
            catch (COMException)
            {
                return 0;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExcelChartUpdateService.CountLinkedCharts");
                return 0;
            }
        }

        /// <summary>Refreshes every Excel-linked chart on a slide; returns how many were updated.</summary>
        public static int UpdateChartsInSlide(Slide slide)
        {
            try
            {
                int updatedCount = 0;

                foreach (Shape shape in slide.Shapes)
                {
                    if (IsLinkedChart(shape) && UpdateChart(shape))
                    {
                        updatedCount++;
                    }
                }

                return updatedCount;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExcelChartUpdateService.UpdateChartsInSlide");
                return 0;
            }
        }

        /// <summary>Refreshes every Excel-linked chart across the given slides; returns the total.</summary>
        public static int UpdateCharts(IEnumerable<Slide> slides)
        {
            int updatedCount = 0;
            if (slides == null) return 0;

            foreach (Slide slide in slides)
            {
                updatedCount += UpdateChartsInSlide(slide);
            }

            return updatedCount;
        }

        private static bool UpdateChart(Shape shape)
        {
            try
            {
                shape.LinkFormat.Update();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ExcelChartUpdateService.UpdateChart");
                return false;
            }
        }
    }
}
