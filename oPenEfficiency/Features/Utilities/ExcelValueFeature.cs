using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace oPenEfficiency.Features.Utilities
{
    /// <summary>
    /// Describes a linked Excel value placed on a slide, serialized into the
    /// shape's AlternativeText (line based, so paths/sheet names are delimiter-safe).
    /// </summary>
    public class ExcelValueLink
    {
        public const string Tag = "oPE_ExcelValue";

        public string Path;   // workbook full name (or file name if unsaved)
        public string Sheet;  // worksheet name
        public string Cell;   // cell address, e.g. $B$3
        public int Decimals = -1; // -1 = show Excel's displayed text as-is
        public bool Thousands;
        public bool Percent;
        public double? Raw;   // raw numeric value (null if the cell is not a number)
        public string Text;   // last displayed text (used in "as Excel" mode)
        public string Id;     // stable identifier for inline links
        public int Start = -1;  // 1-based text start inside a text box for inline links
        public int Length = -1; // current text length for inline links

        public string Build()
        {
            return oPenEfficiency.Core.Services.ExcelValueLinkStatePolicy.BuildTag(
                Path, Sheet, Cell, Decimals, Thousands, Percent, Raw, Id, Start, Length, Text);
        }

        public static bool TryParse(string altText, out ExcelValueLink link)
        {
            link = null;
            if (!oPenEfficiency.Core.Services.ExcelValueLinkStatePolicy.TryParseTag(
                    altText,
                    out string path, out string sheet, out string cell,
                    out int decimals, out bool thousands, out bool percent,
                    out double? raw, out string id, out int start, out int length, out string text))
                return false;

            link = new ExcelValueLink
            {
                Path = path,
                Sheet = sheet,
                Cell = cell,
                Decimals = decimals,
                Thousands = thousands,
                Percent = percent,
                Raw = raw,
                Id = id,
                Start = start,
                Length = length,
                Text = text
            };
            return true;
        }

        /// <summary>Produces the string to display in the shape for the current format.</summary>
        public string FormatValue()
        {
            if (Decimals < 0 || Raw == null)
            {
                return Text ?? "";
            }

            double val = Raw.Value;
            if (Percent) val *= 100.0;

            string nf = Thousands ? "#,##0" : "0";
            if (Decimals > 0) nf += "." + new string('0', Decimals);

            string s = val.ToString(nf, CultureInfo.CurrentCulture);
            return Percent ? s + " %" : s;
        }
    }

    /// <summary>
    /// think-cell-style linked Excel value: inserts the value of the currently
    /// selected Excel cell as a text box at the mouse position in PowerPoint and
    /// keeps it linked so it can be refreshed / number-formatted later
    /// (see the floating toolbar wired in ThisAddIn.ProcessSelectionChange).
    /// </summary>
    [FeatureMetadata(
        Id = "BtnInsertExcelValue",
        Name = "Insert Excel Value",
        Tooltip = "Insert the selected Excel cell value at the mouse position",
        IconData = "",
        Color = "#217346",
        Description = "Inserts the active Excel cell as a refreshable linked text box or inline value in PowerPoint.",
        DetailedHelpText = "### Insert Excel Value\nKeep Excel open and select one cell. In PowerPoint, place the text caret inside a text box to insert the linked value inline; otherwise the command creates a linked text box at the pointer position. Right-click explicitly chooses inline, linked text-box, or unlinked insertion, and can refresh or unlink the selected value. Select a linked result to adjust decimals, thousands separators, and percent formatting. Unlink keeps the displayed text and removes only its Excel connection.",
        RequiredType = OfficeSelectionType.None)]
    public static class InsertExcelValueFeature
    {
        public enum InsertMode
        {
            Automatic,
            InlineAtCursor,
            LinkedTextBox,
            UnlinkedTextBox
        }

        public const string Tag = ExcelValueLink.Tag;

        /// <summary>Shape tag holding all inline linked values inside a single text box.</summary>
        public const string InlineTag = "oPE_InlineValues";
        public const string InlineTemplateTag = "oPE_InlineTemplate";
        public const string InlineStateTag = "oPE_InlineExcelState";

        private const char RecordSeparator = (char)0x1E; // record separator between inline entries
        private const string InlineTokenPrefix = "{{oPE:";
        private const string InlineTokenSuffix = "}}";
        private const string InlineInsertMarkerPrefix = "__oPE_INLINE_INSERT__";

        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, InsertMode.Automatic);
        }

        public static bool Execute(PowerPointManager manager, InsertMode mode)
        {
            try
            {
                var app = manager?.GetApplication();
                if (app == null) return false;

                // 1. Read the selected cell from the running Excel instance.
                Excel.Application xl;
                try
                {
                    xl = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch (COMException)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "No open Excel instance was found. Please select a cell in Excel and try again.");
                    return false;
                }

                Excel.Range cell = null;
                try { cell = xl.ActiveCell; }
                catch (COMException)
                {
                    // Excel has no active cell in the current window context.
                }
                if (cell == null)
                {
                    System.Windows.Forms.MessageBox.Show("No cell is selected in Excel.");
                    return false;
                }

                var ws = (Excel.Worksheet)cell.Worksheet;
                var wb = (Excel.Workbook)ws.Parent;

                var link = new ExcelValueLink
                {
                    Path = SafeString(() => wb.FullName),
                    Sheet = SafeString(() => ws.Name),
                    Cell = SafeString(() => cell.Address),
                    Decimals = -1
                };

                object v2 = null;
                try { v2 = cell.Value2; }
                catch (COMException)
                {
                    // Preserve displayed text even when the raw value is unavailable.
                }
                link.Raw = v2 is double dv ? dv : (double?)null;
                link.Text = SafeString(() => cell.Text as string);
                if (string.IsNullOrEmpty(link.Text) && v2 != null)
                {
                    link.Text = Convert.ToString(v2, CultureInfo.CurrentCulture);
                }

                // 1b. If the cursor is inside a text box, insert the value inline at the
                //     caret (e.g. "revenue is <value>") instead of creating a new box.
                if (mode == InsertMode.Automatic || mode == InsertMode.InlineAtCursor)
                {
                    try
                    {
                        var selection = app.ActiveWindow.Selection;
                        if (TryInsertInline(selection, link)) return true;
                        if (mode == InsertMode.InlineAtCursor)
                        {
                            System.Windows.Forms.MessageBox.Show(
                                "Place the text cursor inside a text box before inserting an inline Excel value.",
                                "Insert Excel Value");
                            return false;
                        }
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "InsertExcelValueFeature.Execute.Inline");
                        if (mode == InsertMode.InlineAtCursor) return false;
                    }
                }

                // 2. Convert the mouse position (screen px) to slide points.
                var slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                if (slide == null) return false;

                GetSlidePointAtMouse(app.ActiveWindow, out float left, out float top);

                // 3. Insert the linked text box.
                var box = slide.Shapes.AddTextbox(
                    MsoTextOrientation.msoTextOrientationHorizontal, left, top, 120f, 20f);
                box.TextFrame.WordWrap = MsoTriState.msoFalse;
                box.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
                box.TextFrame.TextRange.Text = link.FormatValue();
                if (mode != InsertMode.UnlinkedTextBox)
                {
                    box.AlternativeText = link.Build();
                    try { box.Tags.Add(Tag, "1"); }
                    catch (COMException)
                    {
                        // AlternativeText is the authoritative link; the tag is a compatibility hint.
                    }
                }

                box.Select();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "InsertExcelValueFeature.Execute");
                return false;
            }
        }

        public static bool RefreshSelected(PowerPointManager manager)
        {
            try
            {
                var selection = manager?.GetApplication()?.ActiveWindow?.Selection;
                if (selection == null) return false;
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    // The command surface refreshes the selected text box as one scope.
                    // Cursor-targeted refresh remains exclusive to the floating toolbar,
                    // which calls RefreshInlineAtSelection directly.
                    return RefreshInline(manager, GetInlineTargetShape(selection));
                }
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange.Count != 1)
                    return false;
                var shape = selection.ShapeRange[1];
                return RefreshShape(manager, shape) || RefreshInline(manager, shape);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "InsertExcelValueFeature.RefreshSelected");
                return false;
            }
        }

        public static bool UnlinkSelected(PowerPointManager manager)
        {
            try
            {
                var selection = manager?.GetApplication()?.ActiveWindow?.Selection;
                if (selection == null) return false;
                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                    return UnlinkInlineAtSelection(selection);
                if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes || selection.ShapeRange.Count != 1)
                    return false;
                var shape = selection.ShapeRange[1];
                if (ExcelValueLink.TryParse(shape.AlternativeText, out var link))
                {
                    shape.AlternativeText = string.Empty;
                    TryDeleteTag(shape, Tag);
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "InsertExcelValueFeature.UnlinkSelected");
                return false;
            }
        }

        /// <summary>Re-reads the linked cell from Excel and updates the shape text.</summary>
        public static bool RefreshShape(PowerPointManager manager, PowerPoint.Shape shape)
        {
            if (shape == null || !ExcelValueLink.TryParse(SafeString(() => shape.AlternativeText), out var link))
                return false;

            if (!ReadCell(link, out double? raw, out string text, out string error))
            {
                System.Windows.Forms.MessageBox.Show("Refresh failed:\n" + error);
                return false;
            }

            link.Raw = raw;
            if (!string.IsNullOrEmpty(text)) link.Text = text;

            ApplyToShape(shape, link);
            return true;
        }

        /// <summary>Applies a think-cell-style number format and preserves it across refreshes.</summary>
        public static bool SetFormat(PowerPoint.Shape shape, int decimals, bool thousands, bool percent)
        {
            if (shape == null || !ExcelValueLink.TryParse(SafeString(() => shape.AlternativeText), out var link))
                return false;

            if (link.Raw == null)
            {
                System.Windows.Forms.MessageBox.Show(
                    "The linked value is not numeric, so a number format cannot be applied.");
                return false;
            }

            link.Decimals = decimals;
            link.Thousands = thousands;
            link.Percent = percent;
            ApplyToShape(shape, link);
            return true;
        }

        private static void ApplyToShape(PowerPoint.Shape shape, ExcelValueLink link)
        {
            try { shape.TextFrame.TextRange.Text = link.FormatValue(); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "InsertExcelValueFeature.ApplyToShape.Text"); }

            try { shape.AlternativeText = link.Build(); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "InsertExcelValueFeature.ApplyToShape.Alt"); }
        }

        private static bool ReadCell(ExcelValueLink link, out double? raw, out string text, out string error)
        {
            raw = null;
            text = null;
            error = null;

            Excel.Application xl = null;
            bool createdXl = false;
            Excel.Workbook wb = null;
            bool openedWb = false;

            try
            {
                try { xl = (Excel.Application)Marshal.GetActiveObject("Excel.Application"); }
                catch (COMException)
                {
                    xl = new Excel.Application();
                    createdXl = true;
                }

                string fileName = string.IsNullOrEmpty(link.Path) ? null : System.IO.Path.GetFileName(link.Path);
                foreach (Excel.Workbook w in xl.Workbooks)
                {
                    try
                    {
                        if (string.Equals(w.FullName, link.Path, StringComparison.OrdinalIgnoreCase) ||
                            (fileName != null && string.Equals(w.Name, fileName, StringComparison.OrdinalIgnoreCase)))
                        {
                            wb = w;
                            break;
                        }
                    }
                    catch (COMException)
                    {
                        // A workbook can close while the running Excel instance is enumerated.
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "InsertExcelValueFeature.ReadCell.WorkbookProbe");
                    }
                }

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
                var rng = ws.Range[link.Cell];
                object v2 = rng.Value2;
                text = rng.Text as string;
                raw = v2 is double d ? d : (double?)null;
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "InsertExcelValueFeature.ReadCell");
                error = ex.Message;
                return false;
            }
            finally
            {
                try { if (openedWb && wb != null) wb.Close(false); }
                catch (COMException)
                {
                    // The workbook is already closed.
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "InsertExcelValueFeature.ReadCell.CloseWorkbook");
                }

                try { if (createdXl && xl != null && xl.Workbooks.Count == 0) xl.Quit(); }
                catch (COMException)
                {
                    // The owned Excel instance is already shutting down.
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "InsertExcelValueFeature.ReadCell.QuitOwnedExcel");
                }
            }
        }

        // ---- Inline linked values (multiple values inside one text box) ----

        /// <summary>
        /// Inserts a linked value directly into the currently active text selection if possible.
        /// Uses the live TextRange first and resolves the backing shape separately because
        /// ShapeRange access is fragile while PowerPoint is in text-edit mode.
        /// </summary>
        private static bool TryInsertInline(PowerPoint.Selection selection, ExcelValueLink link)
        {
            if (selection == null || selection.Type != PowerPoint.PpSelectionType.ppSelectionText)
            {
                return false;
            }

            PowerPoint.TextRange textRange = null;
            try { textRange = selection.TextRange; }
            catch (COMException)
            {
                // The current selection has no editable text range.
            }
            if (textRange == null)
            {
                return false;
            }

            var textShape = GetInlineTargetShape(selection);
            if (textShape == null)
            {
                return false;
            }

            link.Text = link.FormatValue();
            link.Id = string.IsNullOrEmpty(link.Id) ? Guid.NewGuid().ToString("N") : link.Id;

            var list = ReadInlineLinks(textShape);
            string marker = InlineInsertMarkerPrefix + Guid.NewGuid().ToString("N") + "__";

            PowerPoint.TextRange insertedRange = null;
            try { insertedRange = textRange.InsertAfter(marker); }
            catch (COMException)
            {
                // PowerPoint rejected insertion into the current text selection.
            }
            if (insertedRange == null)
            {
                return false;
            }

            int markerStart = -1;
            try { markerStart = insertedRange.Start; }
            catch (COMException)
            {
                // Continue without position-based replacement when the range is stale.
            }
            if (markerStart > 0)
            {
                foreach (var existingLink in list)
                {
                    if (existingLink.Start >= markerStart)
                    {
                        existingLink.Start += marker.Length;
                    }
                }
            }

            string template = BuildInlineTemplateFromRenderedText(
                SafeString(() => textShape.TextFrame.TextRange.Text),
                list);

            string token = GetInlineToken(link);
            int markerIndex = template.IndexOf(marker, StringComparison.Ordinal);
            if (markerIndex >= 0)
            {
                template = template.Remove(markerIndex, marker.Length).Insert(markerIndex, token);
            }
            else
            {
                template += token;
            }

            list.RemoveAll(l => !string.IsNullOrEmpty(l.Id) && string.Equals(l.Id, link.Id, StringComparison.OrdinalIgnoreCase));
            list.Add(link);
            RenderInlineLinks(textShape, list, template);
            return true;
        }

        /// <summary>
        /// Resolves the shape that owns the current text selection. Prefer ShapeRange when
        /// available, but fall back to the selected child shape because PowerPoint can expose
        /// text edits through either API depending on the editing state.
        /// </summary>
        private static PowerPoint.Shape GetInlineTargetShape(PowerPoint.Selection selection)
        {
            if (selection == null) return null;

            try
            {
                if (selection.ShapeRange != null && selection.ShapeRange.Count >= 1)
                {
                    return selection.ShapeRange[1];
                }
            }
            catch (COMException)
            {
                // ShapeRange is unavailable for some text and child selections.
            }

            try
            {
                if (selection.HasChildShapeRange && selection.ChildShapeRange != null && selection.ChildShapeRange.Count >= 1)
                {
                    return selection.ChildShapeRange[1];
                }
            }
            catch (COMException)
            {
                // ChildShapeRange is optional for the current selection.
            }

            return null;
        }

        private static List<ExcelValueLink> ReadInlineLinks(PowerPoint.Shape shape)
        {
            SynchronizeInlineStateWithVisibleText(shape);
            return ReadInlineState(shape).Links;
        }

        private static string ReadInlineTemplate(PowerPoint.Shape shape)
        {
            SynchronizeInlineStateWithVisibleText(shape);
            return ReadInlineState(shape).Template;
        }

        private static InlineExcelState ReadInlineState(PowerPoint.Shape shape)
        {
            var state = new InlineExcelState();

            string altText = SafeString(() => shape.AlternativeText);
            if (!string.IsNullOrEmpty(altText) && altText.StartsWith(InlineStateTag + "\n", StringComparison.Ordinal))
            {
                try
                {
                    var lines = altText.Replace("\r", "").Split('\n');
                    foreach (var line in lines.Skip(1))
                    {
                        int idx = line.IndexOf('=');
                        if (idx < 0) continue;

                        string key = line.Substring(0, idx);
                        string value = line.Substring(idx + 1);
                        if (key == "template")
                        {
                            state.Template = DecodeBase64(value);
                        }
                        else if (key.StartsWith("link", StringComparison.Ordinal))
                        {
                            var decoded = DecodeBase64(value);
                            if (ExcelValueLink.TryParse(decoded, out var link))
                            {
                                state.Links.Add(link);
                            }
                        }
                    }

                    return state;
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "InsertExcelValueFeature.ReadInlineState.Alt");
                }
            }

            state.Template = ReadLegacyInlineTemplate(shape);
            state.Links = ReadLegacyInlineLinks(shape);
            return state;
        }

        private static List<ExcelValueLink> ReadLegacyInlineLinks(PowerPoint.Shape shape)
        {
            var list = new List<ExcelValueLink>();
            try
            {
                string blob = null;
                for (int i = 1; i <= shape.Tags.Count; i++)
                {
                    if (shape.Tags.Name(i).Equals(InlineTag, StringComparison.OrdinalIgnoreCase))
                    {
                        blob = shape.Tags.Value(i);
                        break;
                    }
                }
                if (string.IsNullOrEmpty(blob)) return list;

                foreach (var record in blob.Split(RecordSeparator))
                {
                    if (ExcelValueLink.TryParse(record, out var l)) list.Add(l);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "InsertExcelValueFeature.ReadInlineLinks");
            }
            return list;
        }

        private static string ReadLegacyInlineTemplate(PowerPoint.Shape shape)
        {
            try
            {
                for (int i = 1; i <= shape.Tags.Count; i++)
                {
                    if (shape.Tags.Name(i).Equals(InlineTemplateTag, StringComparison.OrdinalIgnoreCase))
                    {
                        return shape.Tags.Value(i);
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "InsertExcelValueFeature.ReadInlineTemplate");
            }

            return null;
        }

        private static void WriteInlineLinks(PowerPoint.Shape shape, List<ExcelValueLink> list)
        {
            WriteInlineState(shape, list, ReadInlineTemplate(shape));
        }

        private static void WriteInlineTemplate(PowerPoint.Shape shape, string template)
        {
            WriteInlineState(shape, ReadInlineLinks(shape), template);
        }

        private static void WriteInlineState(PowerPoint.Shape shape, List<ExcelValueLink> links, string template)
        {
            try
            {
                var sb = new StringBuilder();
                sb.Append(InlineStateTag).Append('\n');
                sb.Append("template=").Append(EncodeBase64(template ?? string.Empty)).Append('\n');

                int index = 0;
                foreach (var link in (links ?? new List<ExcelValueLink>()))
                {
                    sb.Append("link").Append(index++).Append('=').Append(EncodeBase64(link.Build())).Append('\n');
                }

                shape.AlternativeText = sb.ToString().TrimEnd('\n');
                shape.Tags.Add(InlineTag, "1");
                shape.Tags.Add(InlineTemplateTag, "1");
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "InsertExcelValueFeature.WriteInlineState");
            }
        }

        private static void SynchronizeInlineStateWithVisibleText(PowerPoint.Shape shape)
        {
            try
            {
                var state = ReadInlineState(shape);
                if (state.Links == null || state.Links.Count == 0)
                {
                    return;
                }

                string rendered = SafeString(() => shape.TextFrame.TextRange.Text);
                var keptLinks = new List<ExcelValueLink>();

                foreach (var link in state.Links.OrderBy(l => l.Start < 1 ? int.MaxValue : l.Start))
                {
                    string display = link.FormatValue();
                    int startIndex = Math.Max(0, link.Start - 1);
                    int length = Math.Max(0, link.Length > -1 ? link.Length : display.Length);

                    if (string.IsNullOrEmpty(display) || startIndex >= rendered.Length || startIndex + length > rendered.Length)
                    {
                        continue;
                    }

                    string current = rendered.Substring(startIndex, length);
                    if (string.Equals(current, display, StringComparison.Ordinal))
                    {
                        keptLinks.Add(link);
                    }
                }

                if (keptLinks.Count == state.Links.Count)
                {
                    return;
                }

                if (keptLinks.Count == 0)
                {
                    ClearInlineState(shape);
                    return;
                }

                string template = BuildInlineTemplateFromRenderedText(rendered, keptLinks);
                WriteInlineState(shape, keptLinks, template);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "InsertExcelValueFeature.SynchronizeInlineStateWithVisibleText");
            }
        }

        /// <summary>Re-reads every inline value in the shape from Excel and replaces the text in place.</summary>
        public static bool RefreshInline(PowerPointManager manager, PowerPoint.Shape shape)
        {
            var list = ReadInlineLinks(shape);
            if (list.Count == 0) return false;

            var template = BuildInlineTemplateFromRenderedText(shape, list);

            int ok = 0;
            string lastError = null;
            foreach (var link in list)
            {
                if (!ReadCell(link, out double? raw, out string excelText, out string error))
                {
                    lastError = error;
                    continue;
                }
                link.Raw = raw;
                link.Text = ComputeDisplay(link, excelText);
                ok++;
            }

            if (ok > 0)
            {
                RenderInlineLinks(shape, list, template);
            }

            if (ok == 0 && lastError != null)
            {
                System.Windows.Forms.MessageBox.Show("Refresh failed:\n" + lastError);
            }
            return ok > 0;
        }

        /// <summary>Applies a number format to all numeric inline values in the shape.</summary>
        public static bool SetFormatInline(PowerPoint.Shape shape, int decimals, bool thousands, bool percent)
        {
            var list = ReadInlineLinks(shape);
            if (list.Count == 0) return false;

            var template = BuildInlineTemplateFromRenderedText(shape, list);

            foreach (var link in list)
            {
                if (link.Raw == null) continue; // skip text values
                link.Decimals = decimals;
                link.Thousands = thousands;
                link.Percent = percent;
                link.Text = link.FormatValue();
            }

            RenderInlineLinks(shape, list, template);
            return true;
        }

        /// <summary>Reads the current format of the first inline value (for toolbar state).</summary>
        public static ExcelValueLink GetFirstInline(PowerPoint.Shape shape)
        {
            var list = ReadInlineLinks(shape);
            return list.Count > 0 ? list[0] : null;
        }

        public static ExcelValueLink GetInlineAtSelection(PowerPoint.Selection selection)
        {
            var shape = GetInlineTargetShape(selection);
            if (shape == null) return null;

            var list = ReadInlineLinks(shape);
            if (list.Count == 0) return null;

            var template = ReadInlineTemplate(shape);
            return FindInlineAtSelection(shape, selection, list, template);
        }

        public static bool RefreshInlineAtSelection(PowerPointManager manager, PowerPoint.Selection selection)
        {
            var shape = GetInlineTargetShape(selection);
            if (shape == null) return false;

            var list = ReadInlineLinks(shape);
            if (list.Count == 0) return false;

            var template = BuildInlineTemplateFromRenderedText(shape, list);

            var targetLink = FindInlineAtSelection(shape, selection, list, template);
            if (targetLink == null) return false;

            if (!ReadCell(targetLink, out double? raw, out string excelText, out string error))
            {
                System.Windows.Forms.MessageBox.Show("Refresh failed:\n" + error);
                return false;
            }

            string oldDisplay = targetLink.FormatValue();
            targetLink.Raw = raw;
            targetLink.Text = ComputeDisplay(targetLink, excelText);
            string newDisplay = targetLink.FormatValue();

            if (!ReplaceInlineAtStoredPosition(shape, list, targetLink, oldDisplay, newDisplay))
            {
                return false;
            }

            WriteInlineState(shape, list, template);
            return true;
        }

        public static bool SetFormatInlineAtSelection(PowerPoint.Selection selection, int decimals, bool thousands, bool percent)
        {
            var shape = GetInlineTargetShape(selection);
            if (shape == null) return false;

            var list = ReadInlineLinks(shape);
            if (list.Count == 0) return false;

            var template = BuildInlineTemplateFromRenderedText(shape, list);

            var targetLink = FindInlineAtSelection(shape, selection, list, template);
            if (targetLink == null || targetLink.Raw == null) return false;

            string oldDisplay = targetLink.FormatValue();
            targetLink.Decimals = decimals;
            targetLink.Thousands = thousands;
            targetLink.Percent = percent;
            string newDisplay = targetLink.FormatValue();
            targetLink.Text = newDisplay;

            if (!ReplaceInlineAtStoredPosition(shape, list, targetLink, oldDisplay, newDisplay))
            {
                return false;
            }

            WriteInlineState(shape, list, template);
            return true;
        }

        public static bool RemoveInlineAtSelection(PowerPoint.Selection selection)
        {
            var shape = GetInlineTargetShape(selection);
            if (shape == null) return false;

            var list = ReadInlineLinks(shape);
            if (list.Count == 0) return false;

            var template = BuildInlineTemplateFromRenderedText(shape, list);
            var targetLink = FindInlineAtSelection(shape, selection, list, template);
            if (targetLink == null) return false;

            string token = GetInlineToken(targetLink);
            int tokenIndex = string.IsNullOrEmpty(template)
                ? -1
                : template.IndexOf(token, StringComparison.Ordinal);

            if (tokenIndex >= 0)
            {
                template = template.Remove(tokenIndex, token.Length);
            }
            else
            {
                template = BuildInlineTemplateWithoutLink(shape, list, targetLink);
            }

            list.RemoveAll(l => !string.IsNullOrEmpty(l.Id) &&
                                string.Equals(l.Id, targetLink.Id, StringComparison.OrdinalIgnoreCase));

            if (list.Count == 0)
            {
                shape.TextFrame.TextRange.Text = template ?? string.Empty;
                ClearInlineState(shape);
                return true;
            }

            RenderInlineLinks(shape, list, template);
            return true;
        }

        public static bool UnlinkInlineAtSelection(PowerPoint.Selection selection)
        {
            var shape = GetInlineTargetShape(selection);
            if (shape == null) return false;

            var list = ReadInlineLinks(shape);
            if (list.Count == 0) return false;

            string rendered = SafeString(() => shape.TextFrame.TextRange.Text);
            var targetLink = FindInlineAtSelection(shape, selection, list, BuildInlineTemplateFromRenderedText(rendered, list));
            if (targetLink == null) return false;

            list.RemoveAll(l => !string.IsNullOrEmpty(l.Id) &&
                                string.Equals(l.Id, targetLink.Id, StringComparison.OrdinalIgnoreCase));

            if (list.Count == 0)
            {
                ClearInlineState(shape);
                return true;
            }

            string template = BuildInlineTemplateFromRenderedText(rendered, list);
            WriteInlineState(shape, list, template);
            return true;
        }

        private static string ComputeDisplay(ExcelValueLink link, string excelText)
        {
            if (link.Decimals < 0)
            {
                return !string.IsNullOrEmpty(excelText)
                    ? excelText
                    : (link.Raw?.ToString(CultureInfo.CurrentCulture) ?? "");
            }
            return link.FormatValue();
        }

        private static string GetInlineToken(ExcelValueLink link)
        {
            return InlineTokenPrefix + (link?.Id ?? string.Empty) + InlineTokenSuffix;
        }

        private static void RenderInlineLinks(PowerPoint.Shape shape, List<ExcelValueLink> links, string template)
        {
            try
            {
                if (shape == null) return;

                if (template == null) template = string.Empty;
                var orderedLinks = links
                    .Where(l => !string.IsNullOrEmpty(l.Id))
                    .OrderBy(l => l.Start < 1 ? int.MaxValue : l.Start)
                    .ToList();

                var sb = new StringBuilder();
                int templateIndex = 0;

                while (templateIndex < template.Length)
                {
                    int tokenStart = template.IndexOf(InlineTokenPrefix, templateIndex, StringComparison.Ordinal);
                    if (tokenStart < 0)
                    {
                        sb.Append(template.Substring(templateIndex));
                        break;
                    }

                    sb.Append(template.Substring(templateIndex, tokenStart - templateIndex));
                    int tokenEnd = template.IndexOf(InlineTokenSuffix, tokenStart, StringComparison.Ordinal);
                    if (tokenEnd < 0)
                    {
                        sb.Append(template.Substring(tokenStart));
                        break;
                    }

                    string token = template.Substring(tokenStart, tokenEnd - tokenStart + InlineTokenSuffix.Length);
                    string id = token.Substring(InlineTokenPrefix.Length, token.Length - InlineTokenPrefix.Length - InlineTokenSuffix.Length);
                    var link = orderedLinks.FirstOrDefault(l => string.Equals(l.Id, id, StringComparison.OrdinalIgnoreCase));
                    if (link != null)
                    {
                        string display = link.FormatValue();
                        link.Text = display;
                        link.Start = sb.Length + 1;
                        link.Length = display.Length;
                        sb.Append(display);
                    }

                    templateIndex = tokenEnd + InlineTokenSuffix.Length;
                }

                shape.TextFrame.TextRange.Text = sb.ToString();
                WriteInlineLinks(shape, orderedLinks);
                WriteInlineTemplate(shape, template);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "InsertExcelValueFeature.RenderInlineLinks");
            }
        }

        private static ExcelValueLink FindInlineAtSelection(
            PowerPoint.Shape shape,
            PowerPoint.Selection selection,
            List<ExcelValueLink> links,
            string template)
        {
            if (shape == null || selection == null || links == null || links.Count == 0) return null;

            try
            {
                var textRange = selection.TextRange;
                if (textRange == null) return null;

                if (!string.IsNullOrEmpty(template))
                {
                    int caret = Math.Max(0, textRange.Start - 1);
                    var linkFromTemplate = FindInlineAtRenderedIndex(template, links, caret);
                    if (linkFromTemplate != null) return linkFromTemplate;

                    if (textRange.Length > 0)
                    {
                        int selectionEnd = Math.Max(caret, caret + textRange.Length - 1);
                        for (int i = caret; i <= selectionEnd; i++)
                        {
                            linkFromTemplate = FindInlineAtRenderedIndex(template, links, i);
                            if (linkFromTemplate != null) return linkFromTemplate;
                        }
                    }
                }

                int start = textRange.Start;
                int length = textRange.Length;
                int end = start + Math.Max(length - 1, 0);

                if (length > 0)
                {
                    var overlap = links.FirstOrDefault(l =>
                        l.Start > 0 &&
                        l.Length > 0 &&
                        l.Start <= end &&
                        (l.Start + l.Length - 1) >= start);
                    if (overlap != null) return overlap;
                }

                int visibleCaret = Math.Max(1, start);
                return links.FirstOrDefault(l =>
                    l.Start > 0 &&
                    l.Length >= 0 &&
                    visibleCaret >= l.Start &&
                    visibleCaret <= (l.Start + Math.Max(l.Length, 1) - 1));
            }
            catch (COMException)
            {
                return null;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "InsertExcelValueFeature.GetInlineAtSelection");
                return null;
            }
        }

        private static ExcelValueLink FindInlineAtRenderedIndex(string template, List<ExcelValueLink> links, int renderedIndex)
        {
            if (string.IsNullOrEmpty(template) || links == null || links.Count == 0) return null;

            renderedIndex = Math.Max(0, renderedIndex);
            int renderedCursor = 0;
            int templateIndex = 0;

            while (templateIndex < template.Length)
            {
                int tokenStart = template.IndexOf(InlineTokenPrefix, templateIndex, StringComparison.Ordinal);
                if (tokenStart < 0)
                {
                    return null;
                }

                int plainLength = tokenStart - templateIndex;
                if (renderedIndex < renderedCursor + plainLength)
                {
                    return null;
                }

                renderedCursor += plainLength;

                int tokenEnd = template.IndexOf(InlineTokenSuffix, tokenStart, StringComparison.Ordinal);
                if (tokenEnd < 0)
                {
                    return null;
                }

                string token = template.Substring(tokenStart, tokenEnd - tokenStart + InlineTokenSuffix.Length);
                string id = token.Substring(InlineTokenPrefix.Length, token.Length - InlineTokenPrefix.Length - InlineTokenSuffix.Length);
                var link = links.FirstOrDefault(l => string.Equals(l.Id, id, StringComparison.OrdinalIgnoreCase));
                int displayLength = Math.Max(link?.FormatValue()?.Length ?? 0, 1);

                if (renderedIndex >= renderedCursor && renderedIndex < renderedCursor + displayLength)
                {
                    return link;
                }

                renderedCursor += displayLength;
                templateIndex = tokenEnd + InlineTokenSuffix.Length;
            }

            return null;
        }

        private static string BuildInlineTemplateFromRenderedText(PowerPoint.Shape shape, List<ExcelValueLink> links)
        {
            return BuildInlineTemplateFromRenderedText(
                SafeString(() => shape.TextFrame.TextRange.Text),
                links);
        }

        private static string BuildInlineTemplateFromRenderedText(string rendered, List<ExcelValueLink> links)
        {
            if (string.IsNullOrEmpty(rendered) || links == null || links.Count == 0)
            {
                return rendered ?? string.Empty;
            }

            var orderedLinks = links
                .Where(l => !string.IsNullOrEmpty(l.Id) && l.Start > 0)
                .OrderByDescending(l => l.Start)
                .ToList();

            foreach (var link in orderedLinks)
            {
                int startIndex = Math.Max(0, link.Start - 1);
                int length = Math.Max(0, link.Length);
                if (startIndex > rendered.Length) continue;
                if (startIndex + length > rendered.Length)
                {
                    length = rendered.Length - startIndex;
                }

                string token = GetInlineToken(link);
                rendered = rendered.Remove(startIndex, length).Insert(startIndex, token);
            }

            return rendered;
        }

        private static string BuildInlineTemplateWithoutLink(
            PowerPoint.Shape shape,
            List<ExcelValueLink> links,
            ExcelValueLink targetLink)
        {
            string rendered = SafeString(() => shape.TextFrame.TextRange.Text);
            if (string.IsNullOrEmpty(rendered) || targetLink == null || targetLink.Start < 1)
            {
                return rendered ?? string.Empty;
            }

            int startIndex = Math.Max(0, targetLink.Start - 1);
            int length = Math.Max(0, targetLink.Length);
            if (startIndex > rendered.Length) startIndex = rendered.Length;
            if (startIndex + length > rendered.Length) length = rendered.Length - startIndex;

            rendered = rendered.Remove(startIndex, length);

            var remainingLinks = links
                .Where(l => !string.IsNullOrEmpty(l.Id) &&
                            !string.Equals(l.Id, targetLink.Id, StringComparison.OrdinalIgnoreCase))
                .Select(l => new ExcelValueLink
                {
                    Path = l.Path,
                    Sheet = l.Sheet,
                    Cell = l.Cell,
                    Decimals = l.Decimals,
                    Thousands = l.Thousands,
                    Percent = l.Percent,
                    Raw = l.Raw,
                    Text = l.Text,
                    Id = l.Id,
                    Start = l.Start > targetLink.Start ? l.Start - length : l.Start,
                    Length = l.Length
                })
                .ToList();

            return BuildInlineTemplateFromRenderedText(rendered, remainingLinks);
        }

        private static void ClearInlineState(PowerPoint.Shape shape)
        {
            try { shape.AlternativeText = string.Empty; }
            catch (COMException)
            {
                // The shape may already have been removed while unlinking.
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "InsertExcelValueFeature.ClearInlineState.AlternativeText");
            }

            TryDeleteTag(shape, InlineTag);
            TryDeleteTag(shape, InlineTemplateTag);
        }

        private static void TryDeleteTag(PowerPoint.Shape shape, string tagName)
        {
            try
            {
                for (int i = shape.Tags.Count; i >= 1; i--)
                {
                    if (shape.Tags.Name(i).Equals(tagName, StringComparison.OrdinalIgnoreCase))
                    {
                        shape.Tags.Delete(tagName);
                        break;
                    }
                }
            }
            catch (COMException)
            {
                // Missing tags and stale shapes are valid cleanup outcomes.
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "InsertExcelValueFeature.TryDeleteTag." + tagName);
            }
        }

        private static bool ReplaceInlineAtStoredPosition(
            PowerPoint.Shape shape,
            List<ExcelValueLink> allLinks,
            ExcelValueLink targetLink,
            string oldDisplay,
            string newDisplay)
        {
            try
            {
                if (shape == null || targetLink == null || targetLink.Start < 1)
                {
                    return false;
                }

                var fullRange = shape.TextFrame.TextRange;
                int oldLength = targetLink.Length > -1
                    ? targetLink.Length
                    : (!string.IsNullOrEmpty(oldDisplay) ? oldDisplay.Length : 0);

                if (oldLength < 0) oldLength = 0;

                var targetRange = fullRange.Characters(targetLink.Start, Math.Max(oldLength, 1));
                targetRange.Text = newDisplay ?? string.Empty;

                int newLength = (newDisplay ?? string.Empty).Length;
                int delta = newLength - oldLength;
                targetLink.Length = newLength;
                targetLink.Text = newDisplay ?? string.Empty;

                if (delta != 0)
                {
                    foreach (var other in allLinks)
                    {
                        if (ReferenceEquals(other, targetLink)) continue;
                        if (other.Start > targetLink.Start)
                        {
                            other.Start += delta;
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "InsertExcelValueFeature.ReplaceInlineAtStoredPosition");
                return false;
            }
        }

        private static string EncodeBase64(string value)
        {
            var bytes = Encoding.UTF8.GetBytes(value ?? string.Empty);
            return Convert.ToBase64String(bytes);
        }

        private static string DecodeBase64(string value)
        {
            if (string.IsNullOrEmpty(value)) return string.Empty;
            var bytes = Convert.FromBase64String(value);
            return Encoding.UTF8.GetString(bytes);
        }

        private sealed class InlineExcelState
        {
            public string Template;
            public List<ExcelValueLink> Links = new List<ExcelValueLink>();
        }

        /// <summary>Inverts PowerPoint's PointsToScreenPixels to map the cursor to slide points.</summary>
        private static void GetSlidePointAtMouse(PowerPoint.DocumentWindow window, out float left, out float top)
        {
            left = 100f;
            top = 100f;
            try
            {
                var mouse = System.Windows.Forms.Cursor.Position;
                double originX = window.PointsToScreenPixelsX(0);
                double originY = window.PointsToScreenPixelsY(0);
                double scaleX = window.PointsToScreenPixelsX(1000) - originX;
                double scaleY = window.PointsToScreenPixelsY(1000) - originY;
                if (Math.Abs(scaleX) > 0.0001) left = (float)((mouse.X - originX) / scaleX * 1000.0);
                if (Math.Abs(scaleY) > 0.0001) top = (float)((mouse.Y - originY) / scaleY * 1000.0);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "InsertExcelValueFeature.GetSlidePointAtMouse");
            }
        }

        private static string SafeString(Func<string> getter)
        {
            try { return getter() ?? ""; }
            catch (COMException) { return ""; }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "InsertExcelValueFeature.SafeString");
                return "";
            }
        }
    }
}
