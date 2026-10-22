using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using WpfApp = System.Windows.Application;
using WinForms = System.Windows.Forms;
using oPenEfficiency.Excel.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using Xl = global::Microsoft.Office.Interop.Excel;

namespace oPenEfficiency.Excel.Features.Data
{
    /// <summary>
    /// Scans the selected range for text entries that look like a date but are not real
    /// calendar dates (e.g. 35.02.2026, 30. Februar 2026, 31 avril 2026), highlights them
    /// and lists them in a floating window where clicking a row jumps to the cell.
    /// Genuine Excel date serials are valid by definition and are ignored.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnValidateDates",
        Name = "Validate Dates",
        Tooltip = "Find invalid date values in the range",
        IconData = "",
        Color = "#EF4444",
        Description = "Highlights cells with impossible date values (for example, 35.02.2026) in the selected range.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Cells)]
    public static class DateValidationFeature
    {
        // Excel "Bad" fill: RGB(255,199,206) expressed as an OLE (BGR) colour.
        private const int BadFillColor = 0xCEC7FF;

        private static DateFindingsWindow _window;

        public static bool Execute(ExcelManager manager)
        {
            try
            {
                var range = manager?.SelectedRange;
                var app = manager?.Application;
                if (range == null || app == null)
                {
                    WinForms.MessageBox.Show("Please select a cell range first.", "Validate Dates",
                        WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                    return false;
                }

                Xl.Range scan = ResolveScanRange(manager) ?? range;
                var sheet = scan.Worksheet as Xl.Worksheet;
                List<DateFinding> findings = ScanRange(scan, true);

                ShowResult(app, sheet, findings);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DateValidationFeature.Execute");
                return false;
            }
        }

        /// <summary>
        /// Runs the production bulk range scan without annotations or UI. This is used by the
        /// isolated Office performance harness and deliberately leaves workbook contents intact.
        /// </summary>
        public static IReadOnlyList<DateFinding> ScanForAutomation(ExcelManager manager)
        {
            Xl.Range scan = ResolveScanRange(manager);
            return scan == null ? new List<DateFinding>() : ScanRange(scan, false);
        }

        private static Xl.Range ResolveScanRange(ExcelManager manager)
        {
            Xl.Range range = manager?.SelectedRange;
            Xl.Application app = manager?.Application;
            if (range == null || app == null) return null;

            // Bound full-column/row selections to the used range for performance.
            try
            {
                Xl.Range used = manager.ActiveSheet?.UsedRange;
                if (used != null)
                {
                    Xl.Range clipped = app.Intersect(range, used) as Xl.Range;
                    if (clipped != null) return clipped;
                }
            }
            catch (COMException)
            {
                // The raw user selection remains the safe fallback when Excel cannot expose
                // UsedRange or Intersect during transient worksheet state.
            }
            catch (InvalidComObjectException)
            {
                // The selection itself remains available even if the optional UsedRange probe
                // became stale between worksheet events.
            }
            return range;
        }

        private static List<DateFinding> ScanRange(Xl.Range scan, bool annotateFindings)
        {
            object raw = scan.Value2;
            int rows = scan.Rows.Count;
            int cols = scan.Columns.Count;
            var findings = new List<DateFinding>();

            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    object value = GetValue(raw, rows, cols, r, c);
                    if (!(value is string text)) continue;

                    string reason = DateValidator.GetInvalidReason(text);
                    if (reason == null) continue;

                    var cell = scan.Cells[r, c] as Xl.Range;
                    if (cell == null) continue;

                    if (annotateFindings)
                    {
                        try { cell.Interior.Color = BadFillColor; }
                        catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "DateValidationFeature.Validate.HighlightCell"); }
                        try { if (cell.Comment == null) cell.AddComment("Invalid date: " + reason); }
                        catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "DateValidationFeature.Validate.AddComment"); }
                    }

                    string address = "?";
                    try { address = cell.Address[false, false]; }
                    catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "DateValidationFeature.Validate.CellAddress"); }
                    findings.Add(new DateFinding { Address = address, Value = text, Reason = reason });
                }
            }
            return findings;
        }

        /// <summary>Reads a cell value from the bulk Value2 result (scalar for a single cell, else a 1-based 2D array).</summary>
        private static object GetValue(object raw, int rows, int cols, int r, int c)
        {
            if (rows == 1 && cols == 1) return raw;
            var arr = raw as object[,];
            return arr?[r, c];
        }

        private static void ShowResult(Xl.Application app, Xl.Worksheet sheet, List<DateFinding> findings)
        {
            if (findings.Count == 0)
            {
                WinForms.MessageBox.Show("No invalid date values found.", "Validate Dates",
                    WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information);
                return;
            }

            EnsureWpfApplication();
            try { _window?.Close(); }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "DateValidationFeature.ShowResults.ClosePreviousWindow"); }

            _window = new DateFindingsWindow(app, sheet, findings);
            _window.Closed += (s, e) => _window = null;
            _window.Show();
            _window.Activate();
        }

        /// <summary>VSTO add-ins have no WPF Application by default; create a singleton so the window has a Dispatcher.</summary>
        private static void EnsureWpfApplication()
        {
            if (WpfApp.Current == null)
            {
                var app = new WpfApp { ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown };
            }
        }
    }
}
