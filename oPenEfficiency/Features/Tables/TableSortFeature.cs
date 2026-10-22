using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnTableSortAZ",
        Name = "Sort A-Z",
        Tooltip = "Sort selected table values A-Z",
        IconData = "M2,11L5.5,2H7.5L11,11H8.8L8,9H5L4.2,11Z M5.8,7H7.2L6.5,5.2Z M3,13H11V15L6,20H11V22H3V20L8,15H3Z M16,2H18V18.6L20.3,16.3L21.7,17.7L17,22.4L12.3,17.7L13.7,16.3L16,18.6Z",
        Color = "#F43F5E",
        Description = "Sorts the selected cells ascending with culture-aware numeric ordering when every compared value is numeric.",
        DetailedHelpText = "### Sort A-Z\n\nSelect at least two cells vertically in the sort column. The primary command sorts only those selected cell texts and keeps destination formatting in place. The **Sort Table Cells** context menu also provides descending and whole-row-text variants. Numeric values use culture-aware numeric order, text uses culture-aware case-insensitive order, and equal values remain stable. Read or write errors trigger rollback of the affected texts and successful sorts create a PowerPoint Undo entry.",
        MinSelection = 0,
        RequiresTable = true)]
    public static class TableSortFeature
    {
        public enum SortType
        {
            AZ_Selection,
            ZA_Selection,
            AZ_WholeRow,
            ZA_WholeRow
        }

        public static bool Execute(PowerPointManager manager) => Execute(manager, SortType.AZ_Selection);

        public static bool Execute(PowerPointManager manager, SortType type)
        {
            if (manager == null) return false;
            bool ascending = type == SortType.AZ_Selection || type == SortType.AZ_WholeRow;
            bool wholeRow = type == SortType.AZ_WholeRow || type == SortType.ZA_WholeRow;

            try
            {
                Shape tableShape;
                Table table;
                if (!TableHelper.GetSelectedTable(manager, out tableShape, out table)) return false;

                int minRow, maxRow, minCol, maxCol;
                if (!TableHelper.GetSelectedCellBounds(table, out minRow, out maxRow, out minCol, out maxCol))
                {
                    ShowInfo("Select at least two cells vertically in the column to sort.");
                    return false;
                }
                if (minRow >= maxRow)
                {
                    ShowInfo("Select at least two rows to sort.");
                    return false;
                }
                if (!wholeRow && minCol != maxCol)
                {
                    ShowInfo("Selected-cell sorting requires one vertical column. Use a whole-row option to move complete row texts.");
                    return false;
                }

                int sortColumn = minCol;
                int firstColumn = wholeRow ? 1 : sortColumn;
                int lastColumn = wholeRow ? table.Columns.Count : sortColumn;
                var original = new Dictionary<string, string>();
                var rows = new List<string[]>();

                for (int row = minRow; row <= maxRow; row++)
                {
                    var cells = new string[lastColumn - firstColumn + 1];
                    for (int column = firstColumn; column <= lastColumn; column++)
                    {
                        string value = table.Cell(row, column).Shape.TextFrame.TextRange.Text;
                        cells[column - firstColumn] = value;
                        original[Key(row, column)] = value;
                    }
                    rows.Add(cells);
                }

                var plannedRows = TableTextPlanner.SortRows(
                    rows,
                    sortColumn - firstColumn,
                    ascending);
                try { manager.GetApplication().StartNewUndoEntry(); } catch { /* Sorting remains valid when undo grouping is unavailable. */ }

                try
                {
                    for (int offset = 0; offset < plannedRows.Count; offset++)
                    {
                        int row = minRow + offset;
                        for (int column = firstColumn; column <= lastColumn; column++)
                            table.Cell(row, column).Shape.TextFrame.TextRange.Text = plannedRows[offset][column - firstColumn];
                    }
                    return true;
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "TableSortFeature.Write");
                    Restore(table, original);
                    System.Windows.Forms.MessageBox.Show("The sort could not be completed. The affected cell texts were restored where PowerPoint allowed it.", "Sort Table Cells", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    return false;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableSortFeature.Execute");
                System.Windows.Forms.MessageBox.Show("The selected table values could not be read, so nothing was sorted.", "Sort Table Cells", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }

        private static void Restore(Table table, Dictionary<string, string> original)
        {
            foreach (var item in original)
            {
                var parts = item.Key.Split(':');
                try { table.Cell(int.Parse(parts[0]), int.Parse(parts[1])).Shape.TextFrame.TextRange.Text = item.Value; }
                catch (Exception ex) { ExceptionLogger.Log(ex, "TableSortFeature.Rollback"); }
            }
        }

        private static string Key(int row, int column) => row + ":" + column;

        private static void ShowInfo(string message)
        {
            System.Windows.Forms.MessageBox.Show(message, "Sort Table Cells", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        }
    }
}
