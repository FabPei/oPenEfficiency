using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using WinForms = System.Windows.Forms;
using Xl = global::Microsoft.Office.Interop.Excel;
using PP = global::Microsoft.Office.Interop.PowerPoint;
using Core = global::Microsoft.Office.Core;
using oPenEfficiency.Excel.Services;
using oPenEfficiency.Services.DataLink;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Excel.Features.PowerPointLink
{
    /// <summary>
    /// Shared bits for the "linked table" feature: the tag stored on the native
    /// PowerPoint table so it can be refreshed later from the source Excel range.
    /// The tag lives in Shape.AlternativeText as
    /// "oPE_TableLink|&lt;workbookPath&gt;|&lt;sheetName&gt;|&lt;rangeAddress&gt;".
    /// </summary>
    internal static class TableLink
    {
        public const string TagPrefix = "oPE_TableLink";

        /// <summary>
        /// Builds the link tag. Format:
        /// "oPE_TableLink|&lt;flag&gt;|&lt;path&gt;|&lt;sheet&gt;|&lt;range&gt;" where flag is
        /// "F" (formatted) or "V" (values only). The flag sits right after the prefix so the
        /// PowerPoint-side parser stays backward compatible with the old flag-less format.
        /// </summary>
        public static string BuildTag(string workbookPath, string sheetName, string rangeAddress, bool formatted)
            => $"{TagPrefix}|{(formatted ? "F" : "V")}|{workbookPath}|{sheetName}|{rangeAddress}";
    }

    /// <summary>
    /// Creates a native PowerPoint table on the active slide from the selected Excel
    /// range and links it (workbook + sheet + range) so it can be refreshed later.
    /// Plain variant: transfers the displayed cell values only, no Excel styling.
    ///
    /// Requires PowerPoint to be running with the target presentation open; the table
    /// is placed on PowerPoint's currently active slide.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnLinkRangeToPptTable",
        Name = "Link as Table",
        Tooltip = "Link the range to PowerPoint as a table",
        IconData = "",
        Color = "#D24726",
        Description = "Creates a native table on the active PowerPoint slide from the selected cell range and stores the data source (file, sheet, range) for later refresh.",
        RequiredType = OfficeSelectionType.None)]
    public static class LinkRangeToPowerPointTableFeature
    {
        /// <summary>Auto-discovery / default entry point: links values only.</summary>
        public static bool Execute(ExcelManager manager) => Execute(manager, false);

        public static bool Execute(ExcelManager manager, bool withFormatting) =>
            ExecuteCore(manager, withFormatting, true);

        public static bool ExecuteForAutomation(ExcelManager manager, bool withFormatting) =>
            ExecuteCore(manager, withFormatting, false);

        private static bool ExecuteCore(ExcelManager manager, bool withFormatting, bool showMessages)
        {
            Xl.XlMousePointer previousCursor = Xl.XlMousePointer.xlDefault;
            Xl.Application cursorApp = null;
            try
            {
                var app = manager?.Application;
                var workbook = manager?.ActiveWorkbook;
                var range = manager?.SelectedRange;

                if (app == null || range == null)
                {
                    ShowMessage(showMessages,
                        "Please select a cell range first.",
                        "Link as Table");
                    return false;
                }

                if (workbook == null)
                {
                    ShowMessage(showMessages,
                        "No active workbook found.",
                        "Link as Table");
                    return false;
                }

                cursorApp = app;
                try { previousCursor = app.Cursor; app.Cursor = Xl.XlMousePointer.xlWait; }
                catch (COMException)
                {
                    // Cursor feedback is optional; linking can continue.
                }

                // A linked table needs a file path so it can be refreshed later.
                // If the workbook has never been saved, prompt the user to save it now.
                if (string.IsNullOrEmpty(workbook.Path))
                {
                    if (!showMessages) return false;
                    try { workbook.Save(); }
                    catch (COMException)
                    {
                        // The user cancelled the Save dialog or Excel rejected the save.
                    }

                    if (string.IsNullOrEmpty(workbook.Path))
                    {
                        // Still unsaved: user cancelled -> abort without linking.
                        return false;
                    }
                }

                int rows = range.Rows.Count;
                int cols = range.Columns.Count;
                if (rows < 1 || cols < 1)
                {
                    return false;
                }

                string workbookPath = workbook.FullName;
                string sheetName = (range.Worksheet as Xl.Worksheet)?.Name ?? string.Empty;
                string rangeAddress = range.Address[
                    false, false, Xl.XlReferenceStyle.xlA1,
                    Type.Missing, Type.Missing];

                // Connect to a running PowerPoint instance (do not start one:
                // the user needs a visible presentation to place the table on).
                // PowerPointConnector attaches via the window even for an unsaved
                // presentation that is not in the Running Object Table.
                bool ambiguousPowerPointInstances;
                PP.Application ppApp = PowerPointConnector.GetRunning(out ambiguousPowerPointInstances);
                if (ppApp == null)
                {
                    ShowMessage(showMessages,
                        ambiguousPowerPointInstances
                            ? "More than one PowerPoint instance is open. Close the non-target instances, then try again. No presentation was changed."
                            : "Please open PowerPoint with the target presentation and show the destination slide.",
                        "Link as Table");
                    return false;
                }

                PP.Slide slide = GetActiveSlide(ppApp);
                if (slide == null)
                {
                    ShowMessage(showMessages,
                        "No active slide found in PowerPoint. Please show a slide.",
                        "Link as Table");
                    return false;
                }

                // Place the table centered on the slide with a sensible size.
                float slideWidth = slide.Master.Width;
                float slideHeight = slide.Master.Height;
                float tableWidth = Math.Min(slideWidth * 0.8f, cols * 120f);
                float tableHeight = Math.Min(slideHeight * 0.7f, rows * 28f);
                float left = (slideWidth - tableWidth) / 2f;
                float top = (slideHeight - tableHeight) / 2f;

                string linkTag = TableLink.BuildTag(workbookPath, sheetName, rangeAddress, withFormatting);
                PP.Shape tableShape = null;

                // Excel and PowerPoint already have a highly optimized native clipboard
                // converter for formatted cell ranges. It creates an editable PowerPoint
                // table in one operation instead of hundreds of cross-process COM calls.
                bool pastedFormattedTable = withFormatting &&
                    TryPasteFormattedTable(range, slide, left, top, tableWidth, tableHeight, out tableShape);

                if (!pastedFormattedTable)
                {
                    float buildLeft = withFormatting ? slideWidth + tableWidth + 50f : left;
                    tableShape = slide.Shapes.AddTable(rows, cols, buildLeft, top, tableWidth, tableHeight);
                }

                // A link created WITH formatting takes its look from Excel, not from a table style:
                // the header row and both bandings are switched off for the pasted frame as well as
                // for the AddTable fallback. The paste used to keep whatever emphasis PowerPoint
                // delivered, and a style fill then covered the transferred Excel fills (FEAT-005
                // DL-FORMAT-CREATE-002, the owner's deck). It runs before the transfer, so the
                // transferred fills stay the last thing written into every cell.
                if (withFormatting) ApplyFormattedTableStyle(tableShape);

                // Creation owns the table geometry: a created table gets the PowerPoint-side
                // column layout (an even division of its own width, exactly like Shapes.AddTable),
                // never Excel's column proportions - column widths are not part of the transfer.
                ApplyPowerPointColumnLayout(tableShape);

                // Values and formatting come from the ONE transfer the PowerPoint-side refresh
                // uses (DataLinkFormatTransfer): an explicit Excel fill keeps its BGR color,
                // Excel "No Fill" arrives as explicit white, and bold, font color, alignment and
                // borders are written by the same code as on refresh. Column widths stay
                // PowerPoint-side on purpose - they are not part of the transfer.
                int transferredCells = DataLinkFormatTransfer.ApplyTable(range, tableShape.Table, withFormatting);
                if (transferredCells == 0)
                {
                    TryDeleteShape(slide.Shapes[tableShape.Name], "EmptyTransferRollback");
                    throw new InvalidOperationException(
                        "No source cell could be transferred into the linked table.");
                }

                string tableShapeName = tableShape.Name;
                tableShape.AlternativeText = linkTag;
                tableShape = slide.Shapes[tableShapeName];
                tableShape.AlternativeText = linkTag;
                tableShape.Left = left;
                tableShape.Top = top;

                // Keep a workbook-local, hidden reference to the source range. Unlike a
                // fill/border marker this survives save/reopen without changing the user's
                // cell formatting, and Excel automatically adjusts it when cells move.
                if (!LinkedRangeRegistry.Register(workbook, range, withFormatting, out string linkId))
                {
                    TryDeleteShape(slide.Shapes[tableShapeName], "RegistrationRollback");
                    throw new InvalidOperationException("The linked range could not be registered in the workbook.");
                }
                try
                {
                    tableShape.Tags.Add("oPE_TableLinkId", linkId);
                }
                catch (Exception ex)
                {
                    LinkedRangeRegistry.Unregister(workbook, linkId);
                    TryDeleteShape(tableShape, "TagRollback");
                    throw new InvalidOperationException(
                        "The linked table could not be tagged; workbook and slide changes were rolled back.",
                        ex);
                }
                UI.ExcelRibbon.InvalidateLinkStatus();

                ApplyOptionalFormatting(() => tableShape.Select(), "SelectCreatedTable");
                ApplyOptionalFormatting(() => ppApp.Activate(), "ActivatePowerPoint");

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "LinkRangeToPowerPointTableFeature.Execute");
                ShowMessage(showMessages,
                    "The table could not be created:\n" + ex.Message,
                    "Link as Table");
                return false;
            }
            finally
            {
                try { if (cursorApp != null) cursorApp.Cursor = previousCursor; }
                catch (COMException)
                {
                    // Excel may already be shutting down.
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "LinkRangeToPowerPointTableFeature.RestoreCursor");
                }
            }
        }

        private static void ShowMessage(bool showMessages, string message, string title)
        {
            if (showMessages) WinForms.MessageBox.Show(message, title);
        }

        private static bool TryPasteFormattedTable(
            Xl.Range source, PP.Slide slide, float left, float top, float width, float height,
            out PP.Shape tableShape)
        {
            tableShape = null;
            System.Windows.Forms.IDataObject previousClipboard = null;
            try
            {
                try { previousClipboard = WinForms.Clipboard.GetDataObject(); }
                catch (ExternalException)
                {
                    // Another process owns the clipboard; continue without a restorable snapshot.
                }
                catch (ThreadStateException)
                {
                    // Clipboard access requires STA; continue with the values-only fallback if paste fails.
                }
                source.Copy();
                WinForms.Application.DoEvents();

                PP.ShapeRange pasted = slide.Shapes.PasteSpecial(PP.PpPasteDataType.ppPasteHTML);
                if (pasted == null || pasted.Count != 1)
                {
                    TryDeleteShapeRange(pasted, "RejectPasteShapeRange");
                    return false;
                }
                PP.Shape candidate = pasted[1];
                if (candidate.HasTable != global::Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    TryDeleteShape(candidate, "RejectNonTablePaste");
                    return false;
                }

                candidate.Left = left;
                candidate.Top = top;
                candidate.Width = width;
                candidate.Height = height;
                tableShape = candidate;
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "LinkRangeToPowerPointTableFeature.TryPasteFormattedTable");
                return false;
            }
            finally
            {
                try
                {
                    if (previousClipboard != null)
                        WinForms.Clipboard.SetDataObject(previousClipboard, true);
                }
                catch (ExternalException ex)
                {
                    ExceptionLogger.Log(ex, "LinkRangeToPowerPointTableFeature.RestoreClipboard");
                }
                catch (ThreadStateException ex)
                {
                    ExceptionLogger.Log(ex, "LinkRangeToPowerPointTableFeature.RestoreClipboard");
                }

                try { source.Application.CutCopyMode = 0; }
                catch (COMException)
                {
                    // Excel may already have cleared copy mode or be shutting down.
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "LinkRangeToPowerPointTableFeature.ClearCopyMode");
                }
            }
        }

        /// <summary>Returns PowerPoint's currently active slide, or null.</summary>
        private static PP.Slide GetActiveSlide(PP.Application ppApp)
        {
            try
            {
                return ppApp.ActiveWindow.View.Slide as PP.Slide;
            }
            catch (COMException)
            {
                try
                {
                    var pres = ppApp.ActivePresentation;
                    if (pres != null && pres.Slides.Count > 0)
                        return pres.Slides[1];
                }
                catch (COMException)
                {
                    // No presentation or slide is currently available.
                }
                return null;
            }
        }

        /// <summary>
        /// Gives a freshly created table the PowerPoint-side column layout: the table's own width
        /// is divided evenly over its columns, which is the layout <c>Shapes.AddTable</c> produces.
        /// Without it the native clipboard paste keeps Excel's column proportions, i.e. imports the
        /// source column widths into the deck. The format contract excludes a width transfer
        /// deliberately (FEAT-005 DL-FORMAT-CREATE-002), and refresh never touches column widths,
        /// so a refreshed table keeps this layout.
        /// </summary>
        private static void ApplyPowerPointColumnLayout(PP.Shape tableShape)
        {
            try
            {
                if (tableShape == null || tableShape.HasTable != global::Microsoft.Office.Core.MsoTriState.msoTrue) return;
                PP.Table table = tableShape.Table;
                int columns = table.Columns.Count;
                if (columns < 1) return;

                float columnWidth = tableShape.Width / columns;
                for (int c = 1; c <= columns; c++)
                {
                    int index = c;
                    ApplyOptionalFormatting(() => table.Columns[index].Width = columnWidth, "Table.ColumnWidth");
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "LinkRangeToPowerPointTableFeature.ApplyPowerPointColumnLayout");
            }
        }

        /// <summary>
        /// Switches the table style's emphasis off (header row, last row, first/last column and both
        /// bandings) so the style cannot paint its own fill over a transferred Excel fill. Only the
        /// formatted route calls it; a values-only link keeps PowerPoint's own table look.
        /// </summary>
        private static void ApplyFormattedTableStyle(PP.Shape tableShape)
        {
            try
            {
                if (tableShape == null || tableShape.HasTable != global::Microsoft.Office.Core.MsoTriState.msoTrue) return;
                PP.Table table = tableShape.Table;
                ApplyOptionalFormatting(() => table.FirstRow = false, "Table.FirstRow");
                ApplyOptionalFormatting(() => table.LastRow = false, "Table.LastRow");
                ApplyOptionalFormatting(() => table.FirstCol = false, "Table.FirstCol");
                ApplyOptionalFormatting(() => table.LastCol = false, "Table.LastCol");
                ApplyOptionalFormatting(() => table.HorizBanding = false, "Table.HorizBanding");
                ApplyOptionalFormatting(() => table.VertBanding = false, "Table.VertBanding");
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "LinkRangeToPowerPointTableFeature.ApplyFormattedTableStyle");
            }
        }

        private static void ApplyOptionalFormatting(Action operation, string context)
        {
            try
            {
                operation();
            }
            catch (COMException)
            {
                // Cross-host formatting support varies by Office version and cell type.
            }
            catch (InvalidCastException)
            {
                // Mixed Excel values can lack a scalar representation.
            }
            catch (FormatException)
            {
                // Invalid optional formatting retains PowerPoint's default.
            }
            catch (OverflowException)
            {
                // Out-of-range optional formatting retains PowerPoint's default.
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "LinkRangeToPowerPointTableFeature." + context);
            }
        }

        private static void TryDeleteShape(PP.Shape shape, string context)
        {
            if (shape == null) return;
            try { shape.Delete(); }
            catch (COMException)
            {
                // The rollback target is already absent.
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "LinkRangeToPowerPointTableFeature." + context);
            }
        }

        private static void TryDeleteShapeRange(PP.ShapeRange shapes, string context)
        {
            if (shapes == null) return;
            try { shapes.Delete(); }
            catch (COMException)
            {
                // The rollback target is already absent.
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "LinkRangeToPowerPointTableFeature." + context);
            }
        }
    }
}
