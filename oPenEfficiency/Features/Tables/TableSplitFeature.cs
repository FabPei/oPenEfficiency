using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>Splits a table into two pieces based on the selected cell.</summary>
    [FeatureMetadata(
        Id = "BtnTableSplit",
        Name = "Split Table",
        Tooltip = "Split table at the selected cell",
        IconData = "M3,3H21V21H3V3M7,5V19H10V5H7M12,5V19H14V5H12M15,5V19H19V5H15Z",
        Color = "#F43F5E",
        Description = "Splits a selected table before the active column, or before the active row from the context menu.",
        DetailedHelpText = "### Split Table\n\nSelect a cell that should become the first cell of the second table. A normal click splits before that column; use **Split Table by Row** in the context menu to split before that row. The command builds both results before replacing the original, so an unsupported merged-cell boundary or another conversion error leaves the source table unchanged. The second table is placed 20 points to the right or below the first.",
        MinSelection = 0,
        RequiresTable = true)]
    public static class TableSplitFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, byColumn: true);
        }

        public static bool Execute(PowerPointManager manager, bool byColumn)
        {
            if (manager == null) return false;

            try
            {
                Shape tableShape;
                Table table;
                if (!TableHelper.GetSelectedTable(manager, out tableShape, out table)) return false;

                int splitIndex = FindSplitIndex(table, byColumn);
                if (splitIndex < 0)
                {
                    MessageBox.Show("Select a table cell at the split position first.", "Split Table", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }

                if (splitIndex <= 1)
                {
                    MessageBox.Show(byColumn
                        ? "Select a cell in column 2 or later. The selected column becomes the first column of the second table."
                        : "Select a cell in row 2 or later. The selected row becomes the first row of the second table.",
                        "Split Table", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }

                Shape firstShape = null;
                Shape secondShape = null;
                try
                {
                    try { manager.GetApplication().StartNewUndoEntry(); } catch { /* Splitting remains valid when undo grouping is unavailable. */ }
                    firstShape = tableShape.Duplicate()[1];
                    secondShape = tableShape.Duplicate()[1];
                    var firstTable = firstShape.Table;
                    var secondTable = secondShape.Table;

                    if (byColumn)
                    {
                        for (int column = firstTable.Columns.Count; column >= splitIndex; column--) firstTable.Columns[column].Delete();
                        for (int column = splitIndex - 1; column >= 1; column--) secondTable.Columns[column].Delete();
                        firstShape.Left = tableShape.Left;
                        firstShape.Top = tableShape.Top;
                        secondShape.Left = firstShape.Left + firstShape.Width + 20;
                        secondShape.Top = tableShape.Top;
                    }
                    else
                    {
                        for (int row = firstTable.Rows.Count; row >= splitIndex; row--) firstTable.Rows[row].Delete();
                        for (int row = splitIndex - 1; row >= 1; row--) secondTable.Rows[row].Delete();
                        firstShape.Left = tableShape.Left;
                        firstShape.Top = tableShape.Top;
                        secondShape.Left = tableShape.Left;
                        secondShape.Top = firstShape.Top + firstShape.Height + 20;
                    }

                    tableShape.Delete();
                    try
                    {
                        firstShape.Select();
                        secondShape.Select(Office.MsoTriState.msoFalse);
                    }
                    catch (Exception ex) { ExceptionLogger.Log(ex, "TableSplitFeature.SelectResults"); }
                    return true;
                }
                catch (Exception ex)
                {
                    TryDelete(firstShape);
                    TryDelete(secondShape);
                    ExceptionLogger.Log(ex, "TableSplitFeature.CreateResults");
                    MessageBox.Show("The table could not be split. The original table was left unchanged. Merged cells that cross the split boundary are not supported.", "Split Table", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableSplitFeature.Execute");
                return false;
            }
        }

        private static int FindSplitIndex(Table table, bool byColumn)
        {
            for (int row = 1; row <= table.Rows.Count; row++)
            {
                for (int column = 1; column <= table.Columns.Count; column++)
                {
                    if (table.Cell(row, column).Selected) return byColumn ? column : row;
                }
            }
            return -1;
        }

        private static void TryDelete(Shape shape)
        {
            if (shape == null) return;
            try { shape.Delete(); } catch { /* Rollback shape may already have been removed by PowerPoint. */ }
        }
    }
}
