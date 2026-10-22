using System;
using System.Globalization;
using System.Linq;
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
            return TryParseDoubleFromCell(text, out double result) ? result : 0;
        }

        public static bool TryParseDoubleFromCell(string text, out double result)
        {
            return oPenEfficiency.Services.FlexibleNumberParser.TryParse(text, out result);
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

        /// <summary>
        /// Resolves the cell that owns an active text cursor. PowerPoint commonly reports
        /// no Cell.Selected flags while the user is editing text inside a table cell.
        /// </summary>
        public static bool TryGetActiveTextCell(PowerPointManager manager, Table table, out int row, out int column)
        {
            row = 0;
            column = 0;
            if (manager == null || table == null) return false;

            try
            {
                var app = manager.GetApplication();
                var selection = app?.ActiveWindow?.Selection;
                if (selection == null || selection.Type != PpSelectionType.ppSelectionText)
                    return false;

                int activeShapeId = 0;
                try
                {
                    dynamic textFrame = selection.TextRange.Parent;
                    Shape cellShape = textFrame?.Parent as Shape;
                    if (cellShape != null) activeShapeId = cellShape.Id;
                }
                catch
                {
                    // Some Office builds do not expose the TextRange parent chain.
                    // The geometry fallback below covers those builds.
                }

                int rowCount = table.Rows.Count;
                int columnCount = table.Columns.Count;
                if (activeShapeId != 0)
                {
                    for (int r = 1; r <= rowCount; r++)
                    {
                        for (int c = 1; c <= columnCount; c++)
                        {
                            if (table.Cell(r, c).Shape.Id == activeShapeId)
                            {
                                row = r;
                                column = c;
                                return true;
                            }
                        }
                    }
                }

                float x;
                float y;
                try
                {
                    var range = selection.TextRange;
                    x = range.BoundLeft + (range.BoundWidth / 2f);
                    y = range.BoundTop + (range.BoundHeight / 2f);
                }
                catch
                {
                    return false;
                }

                const float tolerance = 0.5f;
                for (int r = 1; r <= rowCount; r++)
                {
                    for (int c = 1; c <= columnCount; c++)
                    {
                        var shape = table.Cell(r, c).Shape;
                        if (x >= shape.Left - tolerance &&
                            x <= shape.Left + shape.Width + tolerance &&
                            y >= shape.Top - tolerance &&
                            y <= shape.Top + shape.Height + tolerance)
                        {
                            row = r;
                            column = c;
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableHelper.TryGetActiveTextCell");
            }

            return false;
        }
    }
}
