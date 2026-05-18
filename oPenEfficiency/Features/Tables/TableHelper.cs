using System;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Shared helper for table operations - provides common table selection and cell parsing utilities.
    /// </summary>
    internal static class TableHelper
    {
        /// <summary>
        /// Gets the selected table from the current selection.
        /// Returns false if no table is selected.
        /// </summary>
        public static bool GetSelectedTable(PowerPointManager manager, out Shape tableShape, out Table table)
        {
            tableShape = null;
            table = null;

            if (manager == null) return false;

            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;
                var selection = app.ActiveWindow.Selection;

                if (selection.Type == PpSelectionType.ppSelectionShapes || selection.Type == PpSelectionType.ppSelectionText)
                {
                    var sr = selection.ShapeRange;
                    if (sr != null && sr.Count >= 1 && sr[1].HasTable == Office.MsoTriState.msoTrue)
                    {
                        tableShape = sr[1];
                        table = tableShape.Table;
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableHelper.GetSelectedTable");
            }

            return false;
        }

        /// <summary>
        /// Parses a double value from text, handling various number formats.
        /// </summary>
        public static double ParseDoubleFromCell(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return 0;

            // Extract only digits, minus sign, and one decimal separator
            var sb = new StringBuilder(text.Length);
            bool foundCommaOrDot = false;
            foreach (char ch in text)
            {
                if (char.IsDigit(ch) || ch == '-')
                {
                    sb.Append(ch);
                }
                else if ((ch == '.' || ch == ',') && !foundCommaOrDot)
                {
                    sb.Append(ch);
                    foundCommaOrDot = true;
                }
            }

            string cleaned = sb.ToString();

            // Try parsing with current culture
            if (double.TryParse(cleaned, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.CurrentCulture, out double result))
                return result;

            // Try with invariant culture (dot as decimal separator)
            if (double.TryParse(cleaned.Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out result))
                return result;

            return 0;
        }

        /// <summary>
        /// Gets the bounds of selected cells in a table.
        /// Returns false if no cells are selected.
        /// </summary>
        public static bool GetSelectedCellBounds(Table table, out int minRow, out int maxRow, out int minCol, out int maxCol)
        {
            minRow = int.MaxValue;
            maxRow = int.MinValue;
            minCol = int.MaxValue;
            maxCol = int.MinValue;
            bool anySelected = false;

            // Cache table dimensions to avoid repeated COM calls
            int rowCount = table.Rows.Count;
            int columnCount = table.Columns.Count;

            for (int r = 1; r <= rowCount; r++)
            {
                for (int c = 1; c <= columnCount; c++)
                {
                    var cell = table.Cell(r, c);
                    if (cell.Selected)
                    {
                        anySelected = true;
                        if (r < minRow) minRow = r;
                        if (r > maxRow) maxRow = r;
                        if (c < minCol) minCol = c;
                        if (c > maxCol) maxCol = c;
                    }
                }
            }

            return anySelected;
        }
    }
}
