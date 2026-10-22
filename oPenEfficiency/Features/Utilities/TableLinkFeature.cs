using System;
using System.Runtime.InteropServices;
using WinForms = System.Windows.Forms;
using oPenEfficiency.Services;
using oPenEfficiency.Services.DataLink;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.Features.Utilities
{
    /// <summary>
    /// Metadata for a native PowerPoint table linked to an Excel cell range,
    /// stored in the table shape's AlternativeText as
    /// "oPE_TableLink|&lt;flags&gt;|&lt;workbookPath&gt;|&lt;sheetName&gt;|&lt;rangeAddress&gt;".
    ///
    /// The flags field is a short token of letters: 'F' = linked with formatting,
    /// 'R' = auto-resize the table on refresh; "V" means neither (values only, fixed size).
    /// The range address is always last (e.g. "A1:C4", never contains a pipe) and the path
    /// directly after the flags (Windows paths cannot contain a pipe), so any pipes in
    /// between belong to the sheet name and are rejoined. The flags field is optional so
    /// older flag-less tags ("oPE_TableLink|path|sheet|range") still parse.
    /// </summary>
    public class TableLink
    {
        public const string Tag = "oPE_TableLink";

        public string Path;      // workbook full name
        public string Sheet;     // worksheet name
        public string Range;     // range address, e.g. A1:C4
        public bool Formatted;   // whether the table was linked with Excel formatting
        public bool AutoResize;  // whether refresh resizes the table to match the source range

        public string Build()
        {
            return $"{Tag}|{oPenEfficiency.Core.Services.TableLinkStatePolicy.BuildFlags(Formatted, AutoResize)}|{Path}|{Sheet}|{Range}";
        }

        public static bool TryParse(string altText, out TableLink link)
        {
            link = null;
            if (!oPenEfficiency.Core.Services.TableLinkStatePolicy.TryParseTag(
                    altText, out string path, out string sheet, out string range,
                    out bool formatted, out bool autoResize))
                return false;

            link = new TableLink
            {
                Path = path,
                Range = range,
                Sheet = sheet,
                Formatted = formatted,
                AutoResize = autoResize
            };
            return true;
        }
    }

    /// <summary>
    /// PowerPoint-side operations for a linked table: refresh from the source Excel
    /// range, change the linked range, unlink, and open the source workbook.
    /// The table itself is created on the Excel side (LinkRangeToPowerPointTableFeature).
    /// </summary>
    public static class TableLinkFeature
    {
        public const string Tag = TableLink.Tag;
        public const string LastRefreshTag = "oPE_TableLinkLastRefresh";
        public const string LinkIdTag = "oPE_TableLinkId";

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        private static extern bool IsIconic(IntPtr hWnd);

        [DllImport("oleaut32.dll", EntryPoint = "GetActiveObject", PreserveSig = true)]
        private static extern int GetActiveObjectNative(
            ref Guid rclsid, IntPtr reserved,
            [MarshalAs(UnmanagedType.IUnknown)] out object activeObject);

        private const int SW_RESTORE = 9;

        /// <summary>
        /// Re-reads the linked range (displayed/calculated values) and writes them into
        /// the existing table cells. On any failure the table is left unchanged.
        /// </summary>
        public static bool RefreshTable(PowerPointManager manager, PowerPoint.Shape tableShape)
        {
            if (tableShape == null ||
                !TableLink.TryParse(SafeString(() => tableShape.AlternativeText), out var link))
            {
                return false;
            }

            string linkId = SafeString(() => tableShape.Tags[LinkIdTag]);
            if (!ReadRange(link, linkId, out string[,] values, out DataLinkCellFormat[,] formats, out int rows, out int cols, out string error))
            {
                WinForms.MessageBox.Show(
                    "Refresh failed - the table remains unchanged:\n" + error,
                    "Linked Table");
                return false;
            }

            try
            {
                if (!tableShape.HasTable.Equals(Microsoft.Office.Core.MsoTriState.msoTrue))
                {
                    // Not a table anymore (e.g. user converted it); nothing to write.
                    return false;
                }

                PowerPoint.Table table = tableShape.Table;

                // With auto-resize on, grow/shrink the table to the source dimensions first.
                if (link.AutoResize)
                {
                    try { ResizeTable(table, rows, cols); }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "TableLinkFeature.RefreshTable.Resize");
                        WinForms.MessageBox.Show("Refresh failed while resizing the table:\n" + ex.Message, "Linked Table");
                        return false;
                    }
                }

                int tableRows = table.Rows.Count;
                int tableCols = table.Columns.Count;

                int writeRows = Math.Min(rows, tableRows);
                int writeCols = Math.Min(cols, tableCols);

                bool writeFailed = false;
                int changedCells = 0;
                for (int r = 1; r <= writeRows; r++)
                {
                    for (int c = 1; c <= writeCols; c++)
                    {
                        try
                        {
                            string previous = table.Cell(r, c).Shape.TextFrame.TextRange.Text ?? "";
                            if (!string.Equals(previous, values[r - 1, c - 1] ?? "", StringComparison.Ordinal)) changedCells++;
                            table.Cell(r, c).Shape.TextFrame.TextRange.Text = values[r - 1, c - 1];
                            if (link.Formatted && formats != null)
                                DataLinkFormatTransfer.ApplyCellFormatting(formats[r - 1, c - 1], table.Cell(r, c));
                        }
                        catch (Exception ex) { writeFailed = true; ExceptionLogger.Log(ex, "TableLinkFeature.RefreshTable.Cell"); }
                    }
                }

                if (writeFailed)
                {
                    WinForms.MessageBox.Show("Some table cells could not be updated. Review the table before continuing.", "Linked Table");
                    return false;
                }

                // Only warn about a size mismatch when the user opted out of auto-resize.
                if (!link.AutoResize && (rows != tableRows || cols != tableCols))
                {
                    WinForms.MessageBox.Show(
                        $"The source range now has {rows}x{cols} cells, while the table has {tableRows}x{tableCols}. " +
                        "Only the overlapping cells were updated.\n\n" +
                        "Enable 'Auto-Resize' in the table toolbar to resize the table automatically.",
                        "Linked Table");
                }

                TryBestEffortCom(
                    () => tableShape.Tags.Add(LastRefreshTag, DateTime.Now.ToString("o")),
                    "TableLinkFeature.RefreshTable.LastRefreshTag");
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableLinkFeature.RefreshTable");
                WinForms.MessageBox.Show(
                    "Refresh failed - the table remains unchanged:\n" + ex.Message,
                    "Linked Table");
                return false;
            }
        }

        /// <summary>Prompts for a new range address, stores it and refreshes.</summary>
        public static bool ChangeRange(PowerPointManager manager, PowerPoint.Shape tableShape)
        {
            if (tableShape == null ||
                !TableLink.TryParse(SafeString(() => tableShape.AlternativeText), out var link))
            {
                return false;
            }

            string input = PromptForText(
                "Enter the new cell range (for example, A1:C4):",
                "Change Range",
                link.Range);

            if (string.IsNullOrWhiteSpace(input) || input == link.Range) return false;

            var candidate = new TableLink { Path = link.Path, Sheet = link.Sheet, Range = input.Trim(), Formatted = link.Formatted, AutoResize = link.AutoResize };
            string linkId = SafeString(() => tableShape.Tags[LinkIdTag]);
            if (!ReadRange(candidate, null, out _, out _, out _, out _, out string error))
            {
                WinForms.MessageBox.Show("The range is not valid:\n" + error, "Change Range");
                return false;
            }

            try { tableShape.AlternativeText = candidate.Build(); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "TableLinkFeature.ChangeRange.SetAlt"); return false; }
            SyncRegistryLink(link, candidate, linkId);
            return RefreshTable(manager, tableShape);
        }

        /// <summary>Lets the user repair or replace the workbook, worksheet and range.</summary>
        public static bool RelinkSource(
            PowerPointManager manager,
            PowerPoint.Shape tableShape,
            System.Windows.Window owner = null)
        {
            if (tableShape == null || !TableLink.TryParse(SafeString(() => tableShape.AlternativeText), out var current)) return false;

            var dialog = new oPenEfficiency.UI.RelinkSourceWindow(current.Path, current.Sheet, current.Range);
            if (owner != null)
            {
                dialog.Owner = owner;
            }
            else
            {
                try
                {
                    var app = manager?.GetApplication();
                    if (app != null)
                        new System.Windows.Interop.WindowInteropHelper(dialog).Owner = new IntPtr(app.HWND);
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "TableLinkFeature.RelinkSource.Owner");
                }
            }

            if (dialog.ShowDialog() != true) return false;

            string relinkError;
            if (RelinkSourceForAutomation(
                    manager, tableShape, dialog.WorkbookPath, dialog.Worksheet, dialog.Range, out relinkError))
            {
                return true;
            }
            if (!string.IsNullOrWhiteSpace(relinkError))
            {
                WinForms.MessageBox.Show("The new source is not valid: " + relinkError, "Relink Source");
            }
            return false;
        }

        /// <summary>
        /// Points an existing linked table at another source without collecting the input through the
        /// dialog: workbook, worksheet and range come from the caller, so the whole Source A to
        /// Source B workflow can be exercised without clicking WPF dialogs (DL-RELINK-001).
        ///
        /// The candidate source is read before anything is written, so an invalid workbook, sheet or
        /// range leaves the original link untouched, and a refresh that fails after the payload was
        /// written rolls the payload and the Excel-side registry back. Returns false with
        /// <paramref name="error"/> set whenever the table still points at its previous source.
        /// </summary>
        public static bool RelinkSourceForAutomation(
            PowerPointManager manager,
            PowerPoint.Shape tableShape,
            string workbookPath,
            string worksheet,
            string range,
            out string error)
        {
            error = null;
            if (tableShape == null)
            {
                error = "No linked table was given.";
                return false;
            }

            if (!TableLink.TryParse(SafeString(() => tableShape.AlternativeText), out var current))
            {
                error = "The selected shape does not carry a table link.";
                return false;
            }

            var candidate = new TableLink
            {
                Path = (workbookPath ?? string.Empty).Trim(),
                Sheet = (worksheet ?? string.Empty).Trim(),
                Range = (range ?? string.Empty).Trim(),
                Formatted = current.Formatted,
                AutoResize = current.AutoResize
            };
            if (string.IsNullOrEmpty(candidate.Path) ||
                string.IsNullOrEmpty(candidate.Sheet) ||
                string.IsNullOrEmpty(candidate.Range))
            {
                error = "Workbook, worksheet and range are all required.";
                return false;
            }

            // Validate and read through the ordinary link path first: a bad workbook, sheet or range
            // fails here while the destination still describes the previous source.
            // A wrong workbook, sheet or range is ordinary user input: report it to the caller without
            // stamping a stack trace into the add-in error log (DL-RELINK-001).
            if (!ReadRange(candidate, null, out _, out _, out _, out _, out error, logFailures: false)) return false;

            string linkId = SafeString(() => tableShape.Tags[LinkIdTag]);
            string previousAlternativeText = SafeString(() => tableShape.AlternativeText);
            try { tableShape.AlternativeText = candidate.Build(); }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableLinkFeature.RelinkSourceForAutomation.SetAlt");
                error = "The new source could not be written to the shape.";
                return false;
            }

            SyncRegistryLink(current, candidate, linkId);
            if (RefreshTable(manager, tableShape)) return true;

            // A payload pointing at Source B while the table still shows Source A would be a silent
            // half-switch, so the previous link is restored before the failure is reported.
            try
            {
                tableShape.AlternativeText = previousAlternativeText;
                SyncRegistryLink(candidate, current, linkId);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableLinkFeature.RelinkSourceForAutomation.Rollback");
            }
            error = "The new source could not be applied; the previous link was restored.";
            return false;
        }


        /// <summary>
        /// Flips the auto-resize preference on the linked table and stores it in the tag.
        /// Returns the new state (true = auto-resize on).
        /// </summary>
        public static bool ToggleAutoResize(PowerPoint.Shape tableShape)
        {
            if (tableShape == null ||
                !TableLink.TryParse(SafeString(() => tableShape.AlternativeText), out var link))
            {
                return false;
            }

            link.AutoResize = !link.AutoResize;
            try { tableShape.AlternativeText = link.Build(); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "TableLinkFeature.ToggleAutoResize"); }
            return link.AutoResize;
        }

        /// <summary>Grows or shrinks a table to exactly targetRows × targetCols (min 1×1).</summary>
        private static void ResizeTable(PowerPoint.Table table, int targetRows, int targetCols)
        {
            while (table.Rows.Count < targetRows) table.Rows.Add();
            while (table.Rows.Count > targetRows && table.Rows.Count > 1)
                table.Rows[table.Rows.Count].Delete();

            while (table.Columns.Count < targetCols) table.Columns.Add();
            while (table.Columns.Count > targetCols && table.Columns.Count > 1)
                table.Columns[table.Columns.Count].Delete();
        }

        /// <summary>Removes the link tag so the table is no longer treated as linked.</summary>
        public static bool Unlink(PowerPoint.Shape tableShape)
        {
            if (tableShape == null) return false;
            try
            {
                TableLink.TryParse(SafeString(() => tableShape.AlternativeText), out var oldLink);
                string linkId = SafeString(() => tableShape.Tags[LinkIdTag]);
                if (oldLink != null && !string.IsNullOrEmpty(linkId)) RemoveRegistryLink(oldLink, linkId);
                tableShape.AlternativeText = "";
                TryComProbe(() => tableShape.Tags.Delete(LastRefreshTag));
                TryComProbe(() => tableShape.Tags.Delete(LinkIdTag));
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableLinkFeature.Unlink");
                return false;
            }
        }

        private static void SyncRegistryLink(TableLink oldLink, TableLink newLink, string linkId)
        {
            if (newLink == null || string.IsNullOrEmpty(linkId)) return;
            try
            {
                if (oldLink != null && !string.Equals(oldLink.Path, newLink.Path, StringComparison.OrdinalIgnoreCase))
                    RemoveRegistryLink(oldLink, linkId);
                UpdateRegistryName(newLink, linkId, false);
            }
            catch (Exception ex) { ExceptionLogger.Log(ex, "TableLinkFeature.SyncRegistryLink"); }
        }

        private static void RemoveRegistryLink(TableLink link, string linkId)
        {
            UpdateRegistryName(link, linkId, true);
        }

        private static void UpdateRegistryName(TableLink link, string linkId, bool delete)
        {
            if (link == null || string.IsNullOrEmpty(link.Path) || string.IsNullOrEmpty(linkId)) return;
            Excel.Application xl = null;
            Excel.Workbook wb = null;
            bool createdXl = false;
            bool openedWb = false;
            try
            {
                xl = TryGetActiveExcel();
                if (xl == null) { xl = new Excel.Application(); createdXl = true; }
                wb = FindOpenWorkbook(xl, link.Path);
                if (wb == null)
                {
                    if (!System.IO.File.Exists(link.Path)) return;
                    wb = xl.Workbooks.Open(link.Path);
                    openedWb = true;
                }
                string name = "_oPE_PptTableLink_" + linkId;
                Excel.Name entry = null;
                entry = FindWorkbookName(wb, name);
                if (delete)
                {
                    if (entry != null) entry.Delete();
                }
                else
                {
                    string sheet = (link.Sheet ?? "").Replace("'", "''");
                    string refersTo = "='" + sheet + "'!" + link.Range;
                    if (entry != null) entry.RefersTo = refersTo;
                    else { entry = wb.Names.Add(name, refersTo); entry.Visible = false; }
                }
                wb.Save();
            }
            catch (Exception ex) { ExceptionLogger.Log(ex, "TableLinkFeature.UpdateRegistryName"); }
            finally
            {
                TryBestEffortCom(() => { if (openedWb && wb != null) wb.Close(false); }, "TableLinkFeature.UpdateRegistryName.CloseWorkbook");
                TryBestEffortCom(() => { if (createdXl && xl != null && xl.Workbooks.Count == 0) xl.Quit(); }, "TableLinkFeature.UpdateRegistryName.QuitExcel");
            }
        }

        /// <summary>Opens the source workbook in Excel and selects the linked range.</summary>
        public static bool OpenSource(PowerPoint.Shape tableShape)
        {
            if (tableShape == null ||
                !TableLink.TryParse(SafeString(() => tableShape.AlternativeText), out var link))
            {
                return false;
            }

            if (string.IsNullOrEmpty(link.Path) || !System.IO.File.Exists(link.Path))
            {
                WinForms.MessageBox.Show(
                    "The linked workbook could not be found:\n" + link.Path,
                    "Linked Table");
                return false;
            }

            Excel.Application xl = null;
            bool createdXl = false;
            try
            {
                xl = TryGetActiveExcel();
                if (xl == null) { xl = new Excel.Application(); createdXl = true; }

                xl.Visible = true;

                Excel.Workbook wb = FindOpenWorkbook(xl, link.Path) ?? xl.Workbooks.Open(link.Path);

                var ws = (Excel.Worksheet)wb.Worksheets[link.Sheet];
                TryBestEffortCom(() => wb.Application.Calculate(), "TableLinkFeature.OpenSource.Calculate");
                ws.Activate();
                TryBestEffortCom(() => ws.Range[link.Range].Select(), "TableLinkFeature.OpenSource.SelectRange");

                TryBestEffortCom(() => xl.ActiveWindow.Activate(), "TableLinkFeature.OpenSource.ActivateWindow");

                // Bring the Excel window to the foreground. Activate() alone often loses
                // to the Windows foreground lock, so push it via the window handle.
                try
                {
                    IntPtr hwnd = new IntPtr(xl.Hwnd);
                    if (hwnd != IntPtr.Zero)
                    {
                        if (IsIconic(hwnd)) ShowWindow(hwnd, SW_RESTORE);
                        SetForegroundWindow(hwnd);
                    }
                }
                catch (Exception ex) { ExceptionLogger.Log(ex, "TableLinkFeature.OpenSource.Foreground"); }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableLinkFeature.OpenSource");
                WinForms.MessageBox.Show("The source could not be opened:\n" + ex.Message,
                    "Linked Table");
                // If we started a hidden Excel just for this and it failed, don't leak it.
                TryBestEffortCom(() => { if (createdXl && xl != null && xl.Workbooks.Count == 0) xl.Quit(); }, "TableLinkFeature.OpenSource.QuitExcel");
                return false;
            }
        }

        /// <summary>
        /// Reads the displayed (calculated) values of the linked range. Reuses a running
        /// Excel instance if the workbook is open, otherwise opens it read-only in the
        /// background and closes it again.
        /// </summary>
        private static bool ReadRange(TableLink link, string linkId, out string[,] values, out DataLinkCellFormat[,] formats, out int rows, out int cols, out string error, bool logFailures = true)
        {
            values = null;
            formats = null;
            rows = 0;
            cols = 0;
            error = null;

            Excel.Application xl = null;
            bool createdXl = false;
            Excel.Workbook wb = null;
            bool openedWb = false;

            try
            {
                xl = TryGetActiveExcel();
                if (xl == null) { xl = new Excel.Application(); createdXl = true; }

                wb = FindOpenWorkbook(xl, link.Path);
                if (wb == null)
                {
                    if (string.IsNullOrEmpty(link.Path) || !System.IO.File.Exists(link.Path))
                    {
                        error = "The linked workbook is not open and could not be found.";
                        return false;
                    }
                    wb = xl.Workbooks.Open(link.Path, Type.Missing, true); // ReadOnly
                    openedWb = true;
                }

                var ws = (Excel.Worksheet)wb.Worksheets[link.Sheet];
                Excel.Range rng = null;
                if (!string.IsNullOrEmpty(linkId))
                {
                    TryComProbe(() =>
                    {
                        Excel.Name registryName = FindWorkbookName(wb, "_oPE_PptTableLink_" + linkId);
                        if (registryName != null) rng = registryName.RefersToRange;
                    });
                }
                if (rng == null) rng = ws.Range[link.Range];

                rows = rng.Rows.Count;
                cols = rng.Columns.Count;
                values = new string[rows, cols];
                if (link.Formatted) formats = new DataLinkCellFormat[rows, cols];

                for (int r = 1; r <= rows; r++)
                {
                    for (int c = 1; c <= cols; c++)
                    {
                        try
                        {
                            var cell = (Excel.Range)rng.Cells[r, c];
                            values[r - 1, c - 1] = DataLinkFormatTransfer.ReadDisplayedText(cell);
                            if (link.Formatted) formats[r - 1, c - 1] = DataLinkFormatTransfer.ReadCellFormat(cell);
                        }
                        catch
                        {
                            values[r - 1, c - 1] = string.Empty;
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                if (logFailures) ExceptionLogger.Log(ex, "TableLinkFeature.ReadRange");
                error = ex.Message;
                return false;
            }
            finally
            {
                TryBestEffortCom(() => { if (openedWb && wb != null) wb.Close(false); }, "TableLinkFeature.ReadRange.CloseWorkbook");
                TryBestEffortCom(() => { if (createdXl && xl != null && xl.Workbooks.Count == 0) xl.Quit(); }, "TableLinkFeature.ReadRange.QuitExcel");
            }
        }

        private static Excel.Workbook FindOpenWorkbook(Excel.Application xl, string path)
        {
            if (xl == null || string.IsNullOrEmpty(path)) return null;
            string fileName = System.IO.Path.GetFileName(path);
            foreach (Excel.Workbook w in xl.Workbooks)
            {
                try
                {
                    if (string.Equals(w.FullName, path, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(w.Name, fileName, StringComparison.OrdinalIgnoreCase))
                    {
                        return w;
                    }
                }
                catch (COMException)
                {
                    // A workbook can disappear from the live COM collection during enumeration.
                }
                catch (InvalidComObjectException)
                {
                    // Treat a disconnected workbook proxy as a non-match.
                }
            }
            return null;
        }

        private static Excel.Application TryGetActiveExcel()
        {
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null) return null;
                Guid clsid = excelType.GUID;
                object active;
                return GetActiveObjectNative(ref clsid, IntPtr.Zero, out active) == 0
                    ? active as Excel.Application : null;
            }
            catch { return null; }
        }

        private static Excel.Name FindWorkbookName(Excel.Workbook workbook, string expectedName)
        {
            if (workbook == null || string.IsNullOrEmpty(expectedName)) return null;
            foreach (Excel.Name candidate in workbook.Names)
            {
                try
                {
                    string actual = candidate.Name ?? "";
                    int bang = actual.LastIndexOf('!');
                    if (bang >= 0) actual = actual.Substring(bang + 1);
                    if (string.Equals(actual, expectedName, StringComparison.OrdinalIgnoreCase)) return candidate;
                }
                catch (COMException)
                {
                    // Invalid or hidden workbook names are skipped while probing the collection.
                }
                catch (InvalidComObjectException)
                {
                    // Treat a disconnected name proxy as a non-match.
                }
            }
            return null;
        }

        private static void TryComProbe(Action action)
        {
            try
            {
                action();
            }
            catch (COMException)
            {
                // Missing optional COM state is an expected negative probe result.
            }
            catch (InvalidComObjectException)
            {
                // A disconnected Office proxy is equivalent to missing optional state.
            }
        }

        private static void TryBestEffortCom(Action action, string context)
        {
            try
            {
                action();
            }
            catch (COMException)
            {
                // Best-effort Office activation or cleanup may fail after the COM object changed state.
            }
            catch (InvalidComObjectException)
            {
                // Cleanup is already complete when the Office proxy has disconnected.
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, context);
            }
        }

        /// <summary>Small modal text prompt (PowerPoint has no built-in InputBox).</summary>
        private static string PromptForText(string message, string title, string initial)
        {
            using (var form = new WinForms.Form())
            using (var label = new WinForms.Label())
            using (var textBox = new WinForms.TextBox())
            using (var ok = new WinForms.Button())
            using (var cancel = new WinForms.Button())
            {
                form.Text = title;
                form.FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
                form.StartPosition = WinForms.FormStartPosition.CenterScreen;
                form.MinimizeBox = false;
                form.MaximizeBox = false;
                form.ClientSize = new System.Drawing.Size(360, 120);

                label.SetBounds(12, 12, 336, 20);
                label.Text = message;

                textBox.SetBounds(12, 40, 336, 24);
                textBox.Text = initial ?? "";

                ok.SetBounds(180, 78, 80, 28);
                ok.Text = "OK";
                ok.DialogResult = WinForms.DialogResult.OK;

                cancel.SetBounds(268, 78, 80, 28);
                cancel.Text = "Cancel";
                cancel.DialogResult = WinForms.DialogResult.Cancel;

                form.Controls.AddRange(new WinForms.Control[] { label, textBox, ok, cancel });
                form.AcceptButton = ok;
                form.CancelButton = cancel;

                return form.ShowDialog() == WinForms.DialogResult.OK ? textBox.Text : null;
            }
        }

        private static string SafeString(Func<string> getter)
        {
            try { return getter() ?? ""; }
            catch { return ""; }
        }
    }

    /// <summary>
    /// Sidebar/ribbon entry point that refreshes the currently selected linked table.
    /// Wraps <see cref="TableLinkFeature.RefreshTable"/> so the action is reachable from
    /// the Beta layout without going through the floating selection toolbar.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnRefreshLinkedTable",
        Name = "Refresh Linked Table",
        Tooltip = "Reload the selected linked table from its Excel source",
        IconData = "",
        Color = "#217346",
        Description = "Reloads the linked Excel source of the selected table and writes the latest values into the table cells. The table must have been created via 'Link as Table' in Excel first.",
        DetailedHelpText = "### Refresh Linked Table\nSelect exactly one table previously created through Excel's Link as Table command, then refresh it to load the latest source range. Value-only links preserve PowerPoint formatting; formatted links also apply supported Excel formatting. If Auto-Resize is enabled, the table grows or shrinks to match the source range. Missing files and invalid ranges are reported without removing the existing table content.",
        MinSelection = 1,
        MaxSelection = 1,
        RequiresTable = true,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class RefreshLinkedTableFeature
    {
        public enum LinkedTableAction
        {
            Refresh,
            OpenSource,
            ChangeRange,
            RelinkSource,
            ToggleAutoResize
        }

        public static bool Execute(PowerPointManager manager)
        {
            return ExecuteAction(manager, LinkedTableAction.Refresh);
        }

        public static bool ExecuteAction(PowerPointManager manager, LinkedTableAction action)
        {
            if (manager == null) return false;

            PowerPoint.Shape tableShape = GetSelectedLinkedTable(manager);

            if (tableShape == null)
            {
                WinForms.MessageBox.Show(
                    "Please select a linked table (created via 'Link as Table' in Excel).",
                    "Linked Table");
                return false;
            }

            switch (action)
            {
                case LinkedTableAction.OpenSource:
                    return TableLinkFeature.OpenSource(tableShape);
                case LinkedTableAction.ChangeRange:
                    return TableLinkFeature.ChangeRange(manager, tableShape);
                case LinkedTableAction.RelinkSource:
                    return TableLinkFeature.RelinkSource(manager, tableShape);
                case LinkedTableAction.ToggleAutoResize:
                    try { manager.GetApplication()?.StartNewUndoEntry(); }
                    catch (Exception ex) { ExceptionLogger.Log(ex, "RefreshLinkedTableFeature.ToggleAutoResize.StartUndo"); }
                    TableLinkFeature.ToggleAutoResize(tableShape);
                    return true;
                default:
                    return TableLinkFeature.RefreshTable(manager, tableShape);
            }
        }

        public static bool IsAutoResizeEnabled(PowerPointManager manager)
        {
            var tableShape = GetSelectedLinkedTable(manager);
            return tableShape != null &&
                   TableLink.TryParse(TrySafeAltText(tableShape), out var link) &&
                   link.AutoResize;
        }

        private static PowerPoint.Shape GetSelectedLinkedTable(PowerPointManager manager)
        {
            try
            {
                var shapes = manager?.GetSelectedShapes();
                if (shapes == null || shapes.Count < 1) return null;
                var shape = shapes[1];
                return TableLink.TryParse(TrySafeAltText(shape), out _) ? shape : null;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "RefreshLinkedTableFeature.GetSelectedLinkedTable");
                return null;
            }
        }

        private static string TrySafeAltText(PowerPoint.Shape shape)
        {
            try { return shape.AlternativeText ?? ""; }
            catch { return ""; }
        }
    }
}
