using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;

namespace oPenEfficiency.Services
{
    public static class NativeChartDataSyncManager
    {
        // Storing strong references to the Workbooks prevents the garbage collector from disposing the COM event hooks.
        private static Dictionary<string, Excel.Workbook> _trackedWorkbooks = new Dictionary<string, Excel.Workbook>();

        public static void RegisterHook(Excel.Workbook workbook, Presentation presentation, string chartId)
        {
            if (_trackedWorkbooks.ContainsKey(chartId))
            {
                // Unsubscribe old hook if re-registering
                UnregisterHook(chartId);
            }

            _trackedWorkbooks[chartId] = workbook;

            // Subscribe native Excel application SheetChange event
            workbook.SheetChange += (object sh, Excel.Range target) => HandleSheetChange(workbook, presentation, chartId);
        }

        public static void UnregisterHook(string chartId)
        {
            if (_trackedWorkbooks.TryGetValue(chartId, out var workbook))
            {
                try
                {
                    // Note: We cannot fully unsubscribe the lambda without storing the delegate,
                    // but we can remove the reference to allow cleanup
                    if (workbook != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error releasing workbook: {ex.Message}");
                }
                _trackedWorkbooks.Remove(chartId);
            }
        }

        public static void UnregisterAllHooks()
        {
            foreach (var kvp in _trackedWorkbooks)
            {
                try
                {
                    if (kvp.Value != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(kvp.Value);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error releasing workbook: {ex.Message}");
                }
            }
            _trackedWorkbooks.Clear();
        }

        private static void HandleSheetChange(Excel.Workbook workbook, Presentation presentation, string chartId)
        {
            Excel.Worksheet worksheet = null;
            Excel.Range catCell = null;
            Excel.Range valCell = null;

            try
            {
                worksheet = workbook.Worksheets[1] as Excel.Worksheet;
                if (worksheet == null) return;

                var newRawData = new List<Tuple<string, List<string>>>();

                int row = 2; // Start after header
                while (true)
                {
                    // Clean up previous iteration's COM objects
                    if (catCell != null)
                    {
                        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(catCell); } catch { }
                        catCell = null;
                    }
                    if (valCell != null)
                    {
                        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(valCell); } catch { }
                        valCell = null;
                    }

                    catCell = worksheet.Cells[row, 1] as Excel.Range;
                    valCell = worksheet.Cells[row, 2] as Excel.Range;

                    string cat = catCell?.Value2?.ToString();
                    string val = valCell?.Value2?.ToString();

                    if (string.IsNullOrEmpty(cat) && string.IsNullOrEmpty(val))
                        break; // End of data block

                    newRawData.Add(new Tuple<string, List<string>>(cat ?? "", new List<string> { val ?? "" }));
                    row++;
                }

                if (newRawData.Count > 0)
                {
                    // REDRAW STUB
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Native Chart Sync Error: " + ex.Message);
            }
            finally
            {
                // Always release COM objects
                if (catCell != null)
                {
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(catCell); } catch { }
                }
                if (valCell != null)
                {
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(valCell); } catch { }
                }
                if (worksheet != null)
                {
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet); } catch { }
                }
            }
        }
    }
}
