using System;
using System.Collections.Generic;
using System.Drawing;
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
        Name = "Table Branding",
        Tooltip = "Table branding",
        IconData = "M3,3H21V21H3V3M5,5V10H10V5H5M11,5V10H19V5H11M5,11V19H10V11H5M11,11V19H19V11H11M14,13H17V16H14V13Z",
        Color = "#F43F5E",
        Description = "Save and apply table branding/styles. Opens dialog for style configuration.",
        DetailedHelpText = "### Table Branding\nApplies a predefined corporate color scheme to the selected table, including header row fill, alternating row shading, border colors, and font styles, in one click.",
        MinSelection = 0,
        RequiresTable = true)]
    public static class TableBrandingFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - returns false as Table Branding requires dialog interaction.
        /// Manual switch in MainSidebar.xaml.cs handles opening the TableBrandingDialog.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            // Requires dialog for style configuration
            // Handled by manual switch in MainSidebar.xaml.cs
            return false;
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
            if (manager == null) return false;

            try
            {
                if (!TableHelper.GetSelectedTable(manager, out Shape tableShape, out Table table))
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Please select a table to save the style from.",
                        "Error",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return false;
                }

                // Extract style from table
                var style = ExtractStyleFromTable(table, tableShape);
                style.Name = styleName;
                style.CreatedDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm");

                // Load existing styles and add new one
                var styles = LoadStyles();
                styles.Add(style);
                SaveStyles(styles);

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
        public static bool ApplyTableStyle(PowerPointManager manager, TableStyle style)
        {
            if (manager == null || style == null) return false;

            try
            {
                if (!TableHelper.GetSelectedTable(manager, out Shape tableShape, out Table table))
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Please select a table to apply the style to.",
                        "Error",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return false;
                }

                ApplyStyleToTable(table, tableShape, style);

                System.Windows.Forms.MessageBox.Show(
                    $"Table style '{style.Name}' applied successfully!",
                    "Style Applied",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableBrandingFeature.ApplyTableStyle");
                System.Windows.Forms.MessageBox.Show(
                    $"Error applying table style: {ex.Message}",
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// Extracts style information from a table.
        /// </summary>
        private static TableStyle ExtractStyleFromTable(Table table, Shape tableShape)
        {
            var style = new TableStyle
            {
                HasHeaderRow = table.Application.CommandBars.GetEnabledMso("TableStyleOptions1") ||
                               DetectHeaderRow(table),
                HasAlternatingRows = false
            };

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
                style.RowHeight = table.Rows[1].Height;
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
            catch { }

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
            catch { }
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
                    var line = cell.Line;
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

            // Apply borders to all cells
            for (int i = 1; i <= totalRows; i++)
            {
                for (int j = 1; j <= totalCols; j++)
                {
                    ApplyBorderStyle(table.Cell(i, j).Shape, style);
                }
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

                // Apply outside borders
                if (style.TopBorder != null && style.TopBorder.Visible)
                {
                    var border = borders[PpBorderType.ppBorderTop];
                    border.Visible = Office.MsoTriState.msoTrue;
                    if (style.TopBorder.Color.HasValue) border.ForeColor.RGB = (int)style.TopBorder.Color.Value;
                    if (style.TopBorder.Weight.HasValue) border.Weight = style.TopBorder.Weight.Value;
                    if (style.TopBorder.Style.HasValue) border.Style = style.TopBorder.Style.Value;
                }

                if (style.BottomBorder != null && style.BottomBorder.Visible)
                {
                    var border = borders[PpBorderType.ppBorderBottom];
                    border.Visible = Office.MsoTriState.msoTrue;
                    if (style.BottomBorder.Color.HasValue) border.ForeColor.RGB = (int)style.BottomBorder.Color.Value;
                    if (style.BottomBorder.Weight.HasValue) border.Weight = style.BottomBorder.Weight.Value;
                    if (style.BottomBorder.Style.HasValue) border.Style = style.BottomBorder.Style.Value;
                }

                if (style.LeftBorder != null && style.LeftBorder.Visible)
                {
                    var border = borders[PpBorderType.ppBorderLeft];
                    border.Visible = Office.MsoTriState.msoTrue;
                    if (style.LeftBorder.Color.HasValue) border.ForeColor.RGB = (int)style.LeftBorder.Color.Value;
                    if (style.LeftBorder.Weight.HasValue) border.Weight = style.LeftBorder.Weight.Value;
                    if (style.LeftBorder.Style.HasValue) border.Style = style.LeftBorder.Style.Value;
                }

                if (style.RightBorder != null && style.RightBorder.Visible)
                {
                    var border = borders[PpBorderType.ppBorderRight];
                    border.Visible = Office.MsoTriState.msoTrue;
                    if (style.RightBorder.Color.HasValue) border.ForeColor.RGB = (int)style.RightBorder.Color.Value;
                    if (style.RightBorder.Weight.HasValue) border.Weight = style.RightBorder.Weight.Value;
                    if (style.RightBorder.Style.HasValue) border.Style = style.RightBorder.Style.Value;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableBrandingFeature.ApplyCellBorders");
            }
        }

        /// <summary>
        /// Applies border style to a cell.
        /// </summary>
        private static void ApplyBorderStyle(Shape cell, TableStyle style)
        {
            // Borders are applied through table cell Borders collection, not shape
            // This method is a placeholder - borders are applied in ApplyRowStyle
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
                var styles = JsonService.Load<List<TableStyle>>(StylesFileName);
                return styles ?? new List<TableStyle>();
            }
            catch
            {
                return new List<TableStyle>();
            }
        }

        /// <summary>
        /// Saves table styles to JSON file.
        /// </summary>
        public static void SaveStyles(List<TableStyle> styles)
        {
            try
            {
                JsonService.Save(styles, StylesFileName);
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableBrandingFeature.SaveStyles");
            }
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
