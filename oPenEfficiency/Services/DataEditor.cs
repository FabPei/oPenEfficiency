using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace oPenEfficiency.Services
{
    public class DataEditor : IDisposable
    {
        private Excel.Application _excelApp;
        private Excel.Workbook _workbook;
        private Excel.Worksheet _worksheet;
        private Presentation _presentation;
        private string _currentJsonId;
        private bool _disposed = false;

        public DataEditor(Presentation presentation)
        {
            _presentation = presentation;
        }

        public void OpenDataWindow(string jsonId)
        {
            _currentJsonId = jsonId;

            // 1. Retrieve JSON from CustomXMLParts
            string jsonContent = DataStoreManager.LoadData(_presentation, jsonId);

            // Deserializing placeholder (Assuming we serialize simple Dictionary or List of Tuples)
            // For now, let's use a basic manual parsing or JSON service if available
            // In a real scenario you'd use JsonService.Deserialize<...>(jsonContent)
            var rawData = new List<Tuple<string, List<string>>>();
            if (!string.IsNullOrEmpty(jsonContent))
            {
                try
                {
                    rawData = JsonService.Deserialize<List<Tuple<string, List<string>>>>(jsonContent);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error deserializing data: {ex}");
                    // Continue with empty data - fallback below will handle it
                }
            }

            if (rawData == null || rawData.Count == 0)
            {
                // Fallback mock data if empty
                rawData = new List<Tuple<string, List<string>>>
                {
                    new Tuple<string, List<string>>("Start Year", new List<string> { "100" }),
                    new Tuple<string, List<string>>("Product A", new List<string> { "25" }),
                    new Tuple<string, List<string>>("Product B", new List<string> { "15" }),
                    new Tuple<string, List<string>>("Expenses", new List<string> { "-30" }),
                    new Tuple<string, List<string>>("Subtotal Q1", new List<string> { "e" }),
                    new Tuple<string, List<string>>("Product C", new List<string> { "40" }),
                    new Tuple<string, List<string>>("Taxes", new List<string> { "-15" }),
                    new Tuple<string, List<string>>("End Year", new List<string> { "e" })
                };
            }

            // 2. Instantiate Excel
            _excelApp = new Excel.Application();
            _excelApp.Visible = true;
            // Optionally restrict window size / properties here to make it act like a dialogue
            _excelApp.WindowState = Excel.XlWindowState.xlNormal;

            _workbook = _excelApp.Workbooks.Add("");
            _worksheet = (Excel.Worksheet)_workbook.Worksheets[1];

            // 3. Populate Cells
            _worksheet.Cells[1, 1] = "Category";
            _worksheet.Cells[1, 2] = "Value (Use 'e' for totals)";

            for(int i = 0; i < rawData.Count; i++)
            {
                _worksheet.Cells[i + 2, 1] = rawData[i].Item1;
                _worksheet.Cells[i + 2, 2] = rawData[i].Item2[0];
            }

            // Clean up UI layout
            _worksheet.Columns.AutoFit();

            // 4. Attach Sync Event
            _worksheet.Change += Worksheet_Change;
        }

        private void Worksheet_Change(Excel.Range Target)
        {
            // Sync logic when user changes a cell in Excel
            try
            {
                var newRawData = new List<Tuple<string, List<string>>>();
                // Scan the sheet for data
                int row = 2; // Start after header
                while (true)
                {
                    var catCell = _worksheet.Cells[row, 1] as Excel.Range;
                    var valCell = _worksheet.Cells[row, 2] as Excel.Range;

                    string cat = catCell?.Value2?.ToString();
                    string val = valCell?.Value2?.ToString();

                    if (string.IsNullOrEmpty(cat) && string.IsNullOrEmpty(val))
                        break; // End of data

                    newRawData.Add(new Tuple<string, List<string>>(cat ?? "", new List<string> { val ?? "" }));
                    row++;
                }

                // Reserialize and save
                string newJsonContent = JsonService.Serialize(newRawData);
                DataStoreManager.SaveData(_presentation, _currentJsonId, newJsonContent);

                // REDRAW STUB (Feature specific)
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Error syncing from Excel: " + ex.Message);
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (_disposed)
                return;

            if (disposing)
            {
                // Detach event handlers before cleanup
                if (_worksheet != null)
                {
                    try { _worksheet.Change -= Worksheet_Change; } catch { }
                }

                // Close Excel application
                if (_workbook != null)
                {
                    try { _workbook.Close(false); } catch { }
                }
                if (_excelApp != null)
                {
                    try { _excelApp.Quit(); } catch { }
                }

                // Release COM objects
                if (_worksheet != null)
                {
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(_worksheet); } catch { }
                    _worksheet = null;
                }
                if (_workbook != null)
                {
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(_workbook); } catch { }
                    _workbook = null;
                }
                if (_excelApp != null)
                {
                    try { System.Runtime.InteropServices.Marshal.ReleaseComObject(_excelApp); } catch { }
                    _excelApp = null;
                }
            }

            _disposed = true;
        }

        ~DataEditor()
        {
            Dispose(false);
        }
    }
}
