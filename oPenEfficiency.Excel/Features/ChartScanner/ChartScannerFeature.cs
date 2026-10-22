using System;
using System.Windows;
using oPenEfficiency.Excel.Services;
using oPenEfficiency.Utils;

using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Excel.Features.ChartScanner
{
    /// <summary>
    /// Entry point for the Chart Scanner: opens the interactive digitizer and,
    /// on confirmation, writes the extracted Kategorie/Wert table into the active sheet.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnChartScanner",
        Name = "Chart Scanner",
        Tooltip = "Digitize charts from screenshots into Excel data",
        Color = "#8B5CF6",
        Description = "Digitizes chart images via screen capture/OCR and exports extracted data into Excel.",
        RequiredType = OfficeSelectionType.None)]
    public static class ChartScannerFeature
    {
        public static bool Execute(ExcelManager manager)
        {
            if (manager?.ActiveWorkbook == null)
            {
                MessageBox.Show("Please open a workbook first.", "Chart Scanner",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return false;
            }

            try
            {
                EnsureWpfApplication();

#if WINRT_OCR
                var window = new ChartScannerWindow(new WindowsChartOcr());
#else
                var window = new ChartScannerWindow(new NullChartOcr());
#endif
                bool? ok = window.ShowDialog();
                if (ok != true) return false;

                int rows = ChartScannerExport.Write(manager, window.Result);
                if (rows <= 0)
                {
                    MessageBox.Show("The data could not be written.", "Chart Scanner",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ChartScannerFeature.Execute");
                MessageBox.Show("Chart Scanner error: " + ex.Message, "Chart Scanner",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        /// <summary>
        /// VSTO add-ins have no WPF Application by default; create a singleton so WPF
        /// windows have a Dispatcher and resource scope (mirrors the PowerPoint add-in).
        /// </summary>
        private static void EnsureWpfApplication()
        {
            if (Application.Current == null)
            {
                var app = new Application { ShutdownMode = ShutdownMode.OnExplicitShutdown };
            }
        }
    }
}
