using System;
using oPenEfficiency.Excel.Services;
using oPenEfficiency.Utils;
using Xl = global::Microsoft.Office.Interop.Excel;

namespace oPenEfficiency.Excel.Features.ChartScanner
{
    /// <summary>Writes a scan result into the active worksheet as a clean Kategorie/Wert table.</summary>
    internal static class ChartScannerExport
    {
        /// <summary>
        /// Writes the points starting at the active cell (or A1). Returns the number of rows
        /// written, or -1 on failure. The layout is a two-column table (Kategorie | Wert),
        /// which is directly usable as a chart source and by the Excel→PowerPoint link.
        /// </summary>
        public static int Write(ExcelManager manager, ChartScanResult result)
        {
            try
            {
                var sheet = manager?.ActiveSheet;
                if (sheet == null || result == null || result.Points.Count == 0) return -1;

                Xl.Range anchor = manager.SelectedRange;
                if (anchor != null) anchor = anchor.Cells[1, 1] as Xl.Range;
                if (anchor == null) anchor = sheet.Range["A1"];

                int rows = result.Points.Count + 1; // + header
                var data = new object[rows, 2];
                data[0, 0] = "Kategorie";
                data[0, 1] = "Wert";
                for (int i = 0; i < result.Points.Count; i++)
                {
                    var p = result.Points[i];
                    data[i + 1, 0] = string.IsNullOrWhiteSpace(p.Category) ? (i + 1).ToString() : p.Category;
                    data[i + 1, 1] = Math.Round(p.Value, 6);
                }

                Xl.Range target = anchor.Resize[rows, 2];
                target.Value2 = data;
                try { target.Columns.AutoFit(); } catch { /* Exported values remain valid when Excel cannot auto-fit columns. */ }
                return rows;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ChartScannerExport.Write");
                return -1;
            }
        }
    }
}
