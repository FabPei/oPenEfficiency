using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Interop;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Tables
{
    /// <summary>
    /// Feature: Save and apply table branding/styles (colors, borders, patterns).
    /// Detects header rows and alternating row patterns for intelligent style application.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnTableBranding",
        Name = "Table Style Library",
        SearchAliases = new[] { "Table Branding" },
        Tooltip = "Save and apply reusable table styles",
        IconData = "M3,3H21V21H3V3M5,5V10H10V5H5M11,5V10H19V5H11M5,11V19H10V11H5M11,11V19H19V11H11M14,13H17V16H14V13Z",
        Color = "#F43F5E",
        Description = "Saves reusable table styles and applies their header, body, alternating-row, border, font, and uniform-height settings.",
        DetailedHelpText = "### Table Style Library\n\nSave the selected table's detected header, body, alternating-row, font, four-side border, and uniform row-height styling under a unique name. Applying a style changes formatting but keeps cell text and table structure. Saving an existing name asks before replacement, storage errors are reported, and deleting a style requires confirmation. Header detection uses PowerPoint's table setting when available and otherwise compares the first body rows. Non-uniform source row heights are not flattened into one saved value.",
        MinSelection = 0,
        RequiresTable = true)]
    public static class TableBrandingFeature
    {
        /// <summary>
        /// Opens the style manager for sidebar, search, and Feature Explorer execution.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var dialog = new UI.Dialogs.TableBrandingDialog();
                var app = manager.GetApplication();
                if (app != null) new WindowInteropHelper(dialog).Owner = (IntPtr)app.HWND;
                if (dialog.ShowDialog() != true) return false;
                if (dialog.Action == "Save" && !string.IsNullOrWhiteSpace(dialog.StyleToSave))
                    return SaveTableStyle(manager, dialog.StyleToSave);
                if (dialog.Action == "Apply" && dialog.SelectedStyle != null)
                    return ApplyTableStyle(manager, dialog.SelectedStyle);
                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableBrandingFeature.Execute");
                throw;
            }
        }

        private const string StylesFileName = "table_styles.json";

        /// <summary>
        /// Represents a saved table style/branding.
        /// </summary>
        public class TableStyle
        {
            public string Name { get; set; }
            public string CreatedDate { get; set; }

            // Header row style
            public uint? HeaderFillColor { get; set; }
            public uint? HeaderFontColor { get; set; }
            public string HeaderFontName { get; set; }
            public float? HeaderFontSize { get; set; }
            public bool? HeaderFontBold { get; set; }
            public BorderStyle HeaderBottomBorder { get; set; }

            // Body row styles (for detecting patterns)
            public uint? BodyFillColor { get; set; }
            public uint? BodyFontColor { get; set; }
            public string BodyFontName { get; set; }
            public float? BodyFontSize { get; set; }

            // Alternating row style (every second row)
            public uint? AltRowFillColor { get; set; }
            public uint? AltRowFontColor { get; set; }
            public bool HasAlternatingRows { get; set; }

            // Border styles for all cells
            public BorderStyle TopBorder { get; set; }
            public BorderStyle BottomBorder { get; set; }
            public BorderStyle LeftBorder { get; set; }
            public BorderStyle RightBorder { get; set; }
            public BorderStyle InsideVBorders { get; set; }
            public BorderStyle InsideHBorders { get; set; }

            // Table-level settings
            public float? RowHeight { get; set; }
            public bool HasHeaderRow { get; set; }
        }

        /// <summary>
        /// Border style information.
        /// </summary>
        public class BorderStyle
        {
            public bool Visible { get; set; }
            public uint? Color { get; set; }
            public float? Weight { get; set; }
            public Office.MsoLineStyle? Style { get; set; }
        }

        /// <summary>
        /// Saves the current table's style.
        /// </summary>
        public static bool SaveTableStyle(PowerPointManager manager, string styleName)
        {
            if (manager == null || string.IsNullOrWhiteSpace(styleName)) return false;

            try
            {
                styleName = styleName.Trim();
                if (!TableHelper.GetSelectedTable(manager, out Shape tableShape, out Table table))
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Please select a table to save the style from.",
                        "Error",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return false;
                }

                var styles = LoadStyles();
                var existing = styles.FirstOrDefault(item => string.Equals(item.Name, styleName, StringComparison.OrdinalIgnoreCase));
                if (existing != null)
                {
                    var overwrite = System.Windows.Forms.MessageBox.Show(
                        $"A table style named '{existing.Name}' already exists. Replace it?",
                        "Replace Table Style",
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        System.Windows.Forms.MessageBoxIcon.Question);
                    if (overwrite != System.Windows.Forms.DialogResult.Yes) return false;
                    styles.RemoveAll(item => string.Equals(item.Name, styleName, StringComparison.OrdinalIgnoreCase));
                }

                var style = ExtractStyleFromTable(table, tableShape);
                style.Name = styleName;
                style.CreatedDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm");

                styles.Add(style);
                if (!SaveStyles(styles))
                {
                    System.Windows.Forms.MessageBox.Show("The table style could not be saved. The existing style file was left unchanged.", "Table Style Library", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    return false;
                }

                System.Windows.Forms.MessageBox.Show(
                    $"Table style '{styleName}' saved successfully!",
                    "Style Saved",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableBrandingFeature.SaveTableStyle");
                System.Windows.Forms.MessageBox.Show(
                    $"Error saving table style: {ex.Message}",
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// Applies a saved style to the selected table.
        /// </summary>
        /// <param name="showMessages">
        /// False selects the message-box-free programmatic path (acceptance seam); the table
        /// styling behavior and return value are identical either way. Default is true, so the
        /// dialog-driven user path is unchanged.
        /// </param>
        public static bool ApplyTableStyle(PowerPointManager manager, TableStyle style, bool showMessages = true)
        {
            if (manager == null || style == null) return false;

            try
            {
                if (!TableHelper.GetSelectedTable(manager, out Shape tableShape, out Table table))
                {
                    if (showMessages)
                    {
                        System.Windows.Forms.MessageBox.Show(
                            "Please select a table to apply the style to.",
                            "Error",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Warning);
                    }
                    return false;
                }

                try { manager.GetApplication().StartNewUndoEntry(); } catch { /* Branding remains valid when undo grouping is unavailable. */ }
                ApplyStyleToTable(table, tableShape, style);

                if (showMessages)
                {
                    System.Windows.Forms.MessageBox.Show(
                        $"Table style '{style.Name}' applied successfully!",
                        "Style Applied",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableBrandingFeature.ApplyTableStyle");
                if (showMessages)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "The table style could not be applied completely. Use PowerPoint Undo to restore the previous formatting.",
                        "Error",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Error);
                }
                return false;
            }
        }

        /// <summary>
        /// Extracts style information from a table.
        /// </summary>
        private static TableStyle ExtractStyleFromTable(Table table, Shape tableShape)
        {
            bool hasHeaderRow;
            try { hasHeaderRow = table.FirstRow; }
            catch { hasHeaderRow = DetectHeaderRow(table); }
            if (!hasHeaderRow) hasHeaderRow = DetectHeaderRow(table);

            var style = new TableStyle { HasHeaderRow = hasHeaderRow, HasAlternatingRows = false };

            int totalRows = table.Rows.Count;
            if (totalRows == 0) return style;

            // Analyze header row (row 1)
            if (style.HasHeaderRow && totalRows >= 1)
            {
                ExtractRowStyle(table, 1, style, isHeader: true);
            }

            // Analyze body rows for patterns
            if (totalRows > 1)
            {
                AnalyzeBodyRows(table, style, totalRows);
            }

            // Extract border styles from a representative cell
            if (table.Rows.Count >= 1 && table.Columns.Count >= 1)
            {
                ExtractBorderStyle(table.Cell(1, 1), style);
            }

            // Extract row height
            if (table.Rows.Count >= 1)
            {
                style.RowHeight = GetUniformRowHeight(table);
            }

            return style;
        }

        /// <summary>
        /// Detects if the first row appears to be a header row.
        /// </summary>
        private static bool DetectHeaderRow(Table table)
        {
            if (table.Rows.Count < 2) return false;

            var row1 = table.Rows[1];
            var row2 = table.Rows[2];

            // Check if row 1 has different formatting than row 2
            try
            {
                var cell1 = row1.Cells[1].Shape;
                var cell2 = row2.Cells[1].Shape;

                // Different fill color?
                if (cell1.Fill.Visible == Office.MsoTriState.msoTrue &&
                    (cell2.Fill.Visible == Office.MsoTriState.msoFalse ||
                     cell1.Fill.ForeColor.RGB != cell2.Fill.ForeColor.RGB))
                {
                    return true;
                }

                // Bold text in row 1?
                if (cell1.HasTextFrame == Office.MsoTriState.msoTrue &&
                    cell2.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    if (cell1.TextFrame.TextRange.Font.Bold == Office.MsoTriState.msoTrue &&
                        cell2.TextFrame.TextRange.Font.Bold == Office.MsoTriState.msoFalse)
                    {
                        return true;
                    }
                }
            }
            catch { /* Unsupported optional table text formatting retains the existing cell style. */ }

            return false;
        }

        /// <summary>
        /// Analyzes body rows to detect alternating row patterns.
        /// </summary>
        private static void AnalyzeBodyRows(Table table, TableStyle style, int totalRows)
        {
            int startRow = style.HasHeaderRow ? 2 : 1;
            if (totalRows < startRow + 1) return;

            var evenColors = new List<uint>();
            var oddColors = new List<uint>();

            try
            {
                // Sample colors from rows
                for (int i = startRow; i <= totalRows && i < startRow + 6; i++)
                {
                    var cell = table.Rows[i].Cells[1].Shape;
                    if (cell.Fill.Visible == Office.MsoTriState.msoTrue)
                    {
                        uint color = (uint)cell.Fill.ForeColor.RGB;
                        if ((i - startRow) % 2 == 0)
                            oddColors.Add(color);
                        else
                            evenColors.Add(color);
                    }
                }

                // Check for alternating pattern
                if (oddColors.Count > 0 && evenColors.Count > 0)
                {
                    var oddMode = GetMostCommon(oddColors);
                    var evenMode = GetMostCommon(evenColors);

                    if (oddMode != evenMode && oddColors.Count >= 2 && evenColors.Count >= 2)
                    {
                        style.HasAlternatingRows = true;
                        style.BodyFillColor = oddMode;
                        style.AltRowFillColor = evenMode;
                    }
                    else
                    {
                        style.BodyFillColor = oddMode;
                    }
                }
                else if (oddColors.Count > 0)
                {
                    style.BodyFillColor = GetMostCommon(oddColors);
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableBrandingFeature.AnalyzeBodyRows");
            }

            // Extract font colors
            try
            {
                int sampleRow = startRow;
                var cell = table.Rows[sampleRow].Cells[1].Shape;
                if (cell.HasTextFrame == Office.MsoTriState.msoTrue && cell.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    style.BodyFontColor = (uint)cell.TextFrame.TextRange.Font.Color.RGB;
                    style.BodyFontName = cell.TextFrame.TextRange.Font.Name;
                    style.BodyFontSize = cell.TextFrame.TextRange.Font.Size;
                }
            }
            catch { /* Unsupported optional table border formatting retains the existing border. */ }
        }

        /// <summary>
        /// Extracts style from a specific row.
        /// </summary>
        private static void ExtractRowStyle(Table table, int rowNumber, TableStyle style, bool isHeader)
        {
            try
            {
                var cell = table.Rows[rowNumber].Cells[1].Shape;

                // Fill color
                if (cell.Fill.Visible == Office.MsoTriState.msoTrue)
                {
                    uint color = (uint)cell.Fill.ForeColor.RGB;
                    if (isHeader)
                        style.HeaderFillColor = color;
                    else if (!isHeader && !style.HasAlternatingRows)
                        style.BodyFillColor = color;
                }

                // Font properties
                if (cell.HasTextFrame == Office.MsoTriState.msoTrue && cell.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    var font = cell.TextFrame.TextRange.Font;
                    if (isHeader)
                    {
                        style.HeaderFontColor = (uint?)font.Color.RGB;
                        style.HeaderFontName = font.Name;
                        style.HeaderFontSize = font.Size;
                        style.HeaderFontBold = font.Bold == Office.MsoTriState.msoTrue;
                    }
                    else
                    {
                        style.BodyFontColor = (uint?)font.Color.RGB;
                        style.BodyFontName = font.Name;
                        style.BodyFontSize = font.Size;
                    }
                }

                // Header bottom border (special case for header row)
                if (isHeader)
                {
                    var line = table.Cell(rowNumber, 1).Borders[PpBorderType.ppBorderBottom];
                    style.HeaderBottomBorder = new BorderStyle
                    {
                        Visible = line.Visible == Office.MsoTriState.msoTrue,
                        Color = (uint?)line.ForeColor.RGB,
                        Weight = line.Weight,
                        Style = line.Style
                    };
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableBrandingFeature.ExtractRowStyle");
            }
        }

        /// <summary>
        /// Extracts border style from a cell.
        /// </summary>
        private static void ExtractBorderStyle(Cell cell, TableStyle style)
        {
            try
            {
                // Access borders from the cell
                var borders = cell.Borders;

                // For table cells, we check the border properties
                style.TopBorder = new BorderStyle
                {
                    Visible = borders[PpBorderType.ppBorderTop].Visible == Office.MsoTriState.msoTrue,
                    Color = (uint)borders[PpBorderType.ppBorderTop].ForeColor.RGB,
                    Weight = borders[PpBorderType.ppBorderTop].Weight,
                    Style = borders[PpBorderType.ppBorderTop].Style
                };

                style.BottomBorder = new BorderStyle
                {
                    Visible = borders[PpBorderType.ppBorderBottom].Visible == Office.MsoTriState.msoTrue,
                    Color = (uint)borders[PpBorderType.ppBorderBottom].ForeColor.RGB,
                    Weight = borders[PpBorderType.ppBorderBottom].Weight,
                    Style = borders[PpBorderType.ppBorderBottom].Style
                };

                style.LeftBorder = new BorderStyle
                {
                    Visible = borders[PpBorderType.ppBorderLeft].Visible == Office.MsoTriState.msoTrue,
                    Color = (uint)borders[PpBorderType.ppBorderLeft].ForeColor.RGB,
                    Weight = borders[PpBorderType.ppBorderLeft].Weight,
                    Style = borders[PpBorderType.ppBorderLeft].Style
                };

                style.RightBorder = new BorderStyle
                {
                    Visible = borders[PpBorderType.ppBorderRight].Visible == Office.MsoTriState.msoTrue,
                    Color = (uint)borders[PpBorderType.ppBorderRight].ForeColor.RGB,
                    Weight = borders[PpBorderType.ppBorderRight].Weight,
                    Style = borders[PpBorderType.ppBorderRight].Style
                };
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableBrandingFeature.ExtractBorderStyle");
            }
        }

        /// <summary>
        /// Applies a style to a table.
        /// </summary>
        private static void ApplyStyleToTable(Table table, Shape tableShape, TableStyle style)
        {
            int totalRows = table.Rows.Count;
            int totalCols = table.Columns.Count;

            // Apply header row style
            if (style.HasHeaderRow && totalRows >= 1)
            {
                ApplyRowStyle(table, 1, style, isHeader: true);
            }

            // Apply body row styles (with alternating pattern if detected)
            int startRow = style.HasHeaderRow ? 2 : 1;
            for (int i = startRow; i <= totalRows; i++)
            {
                bool isAltRow = style.HasAlternatingRows && ((i - startRow) % 2 == 1);
                ApplyRowStyle(table, i, style, isHeader: false, isAltRow: isAltRow);
            }

            // Apply row height if specified
            if (style.RowHeight.HasValue)
            {
                for (int i = 1; i <= totalRows; i++)
                {
                    table.Rows[i].Height = style.RowHeight.Value;
                }
            }
        }

        /// <summary>
        /// Applies row style to a specific row.
        /// </summary>
        private static void ApplyRowStyle(Table table, int rowNumber, TableStyle style, bool isHeader, bool isAltRow = false)
        {
            try
            {
                int totalCols = table.Columns.Count;

                uint? fillColor = isHeader ? style.HeaderFillColor :
                                  (isAltRow ? style.AltRowFillColor : style.BodyFillColor);
                uint? fontColor = isHeader ? style.HeaderFontColor :
                                  (isAltRow ? style.AltRowFontColor : style.BodyFontColor);
                string fontName = isHeader ? style.HeaderFontName : style.BodyFontName;
                float? fontSize = isHeader ? style.HeaderFontSize : style.BodyFontSize;
                bool? fontBold = isHeader ? style.HeaderFontBold : null;

                for (int col = 1; col <= totalCols; col++)
                {
                    var cell = table.Rows[rowNumber].Cells[col];

                    // Apply fill color
                    if (fillColor.HasValue)
                    {
                        cell.Shape.Fill.Visible = Office.MsoTriState.msoTrue;
                        cell.Shape.Fill.ForeColor.RGB = (int)fillColor.Value;
                    }

                    // Apply font properties
                    if (cell.Shape.HasTextFrame == Office.MsoTriState.msoTrue && cell.Shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        var font = cell.Shape.TextFrame.TextRange.Font;
                        if (fontColor.HasValue)
                            font.Color.RGB = (int)fontColor.Value;
                        if (!string.IsNullOrEmpty(fontName))
                            font.Name = fontName;
                        if (fontSize.HasValue)
                            font.Size = fontSize.Value;
                        if (fontBold.HasValue)
                            font.Bold = fontBold.Value ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                    }

                    // Apply borders to this cell
                    ApplyCellBorders(cell, style);
                }

                // Apply header bottom border (special case - emphasized border under header row)
                if (isHeader && style.HeaderBottomBorder != null && style.HeaderBottomBorder.Visible)
                {
                    for (int col = 1; col <= totalCols; col++)
                    {
                        var cell = table.Rows[rowNumber].Cells[col];
                        // Apply bottom border to the cell
                        var cellBorders = cell.Borders;
                        var bottomBorder = cellBorders[PpBorderType.ppBorderBottom];
                        bottomBorder.Visible = Office.MsoTriState.msoTrue;
                        bottomBorder.ForeColor.RGB = (int)(style.HeaderBottomBorder.Color ?? 0x000000u);
                        bottomBorder.Weight = style.HeaderBottomBorder.Weight ?? 1f;
                        if (style.HeaderBottomBorder.Style.HasValue)
                            bottomBorder.Style = style.HeaderBottomBorder.Style.Value;
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableBrandingFeature.ApplyRowStyle");
                throw;
            }
        }

        /// <summary>
        /// Applies border styles to a table cell.
        /// </summary>
        private static void ApplyCellBorders(Cell cell, TableStyle style)
        {
            try
            {
                var borders = cell.Borders;

                ApplySingleBorder(borders, PpBorderType.ppBorderTop, style.TopBorder);
                ApplySingleBorder(borders, PpBorderType.ppBorderBottom, style.BottomBorder);
                ApplySingleBorder(borders, PpBorderType.ppBorderLeft, style.LeftBorder);
                ApplySingleBorder(borders, PpBorderType.ppBorderRight, style.RightBorder);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableBrandingFeature.ApplyCellBorders");
                throw;
            }
        }

        /// <summary>
        /// Applies a single border to a cell.
        /// </summary>
        private static void ApplySingleBorder(Borders borders, PpBorderType borderType, BorderStyle borderStyle)
        {
            if (borderStyle == null) return;

            var border = borders[borderType];
            border.Visible = borderStyle.Visible ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;

            if (borderStyle.Visible && borderStyle.Color.HasValue)
            {
                border.ForeColor.RGB = (int)borderStyle.Color.Value;
            }

            if (borderStyle.Weight.HasValue)
            {
                border.Weight = borderStyle.Weight.Value;
            }

            if (borderStyle.Style.HasValue)
            {
                border.Style = borderStyle.Style.Value;
            }
        }

        /// <summary>
        /// Loads saved table styles from JSON file.
        /// </summary>
        public static List<TableStyle> LoadStyles()
        {
            try
            {
                var styles = JsonService.Load<List<TableStyle>>(StylesFileName) ?? new List<TableStyle>();
                return styles
                    .Where(style => style != null && !string.IsNullOrWhiteSpace(style.Name))
                    .GroupBy(style => style.Name, StringComparer.OrdinalIgnoreCase)
                    .Select(group => group.Last())
                    .ToList();
            }
            catch
            {
                return new List<TableStyle>();
            }
        }

        /// <summary>
        /// Saves table styles to JSON file.
        /// </summary>
        public static bool SaveStyles(List<TableStyle> styles)
        {
            try
            {
                return JsonService.TrySave(styles ?? new List<TableStyle>(), StylesFileName);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableBrandingFeature.SaveStyles");
                return false;
            }
        }

        private static float? GetUniformRowHeight(Table table)
        {
            if (table.Rows.Count == 0) return null;
            float first = table.Rows[1].Height;
            for (int row = 2; row <= table.Rows.Count; row++)
            {
                if (Math.Abs(table.Rows[row].Height - first) > 0.1f) return null;
            }
            return first;
        }

        /// <summary>
        /// Gets the most common color from a list.
        /// </summary>
        private static uint GetMostCommon(List<uint> colors)
        {
            if (colors.Count == 0) return 0;

            var groups = new Dictionary<uint, int>();
            foreach (var c in colors)
            {
                if (groups.ContainsKey(c))
                    groups[c]++;
                else
                    groups[c] = 1;
            }

            uint mostCommon = colors[0];
            int maxCount = 0;
            foreach (var kvp in groups)
            {
                if (kvp.Value > maxCount)
                {
                    maxCount = kvp.Value;
                    mostCommon = kvp.Key;
                }
            }

            return mostCommon;
        }
    }
}
