using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Insert a new row into a table relative to the selection.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnTableRowInsertion",
        Name = "Insert Table Row",
        SearchAliases = new[] { "Insert Row" },
        Tooltip = "Insert a table row at the selected position",
        IconData = "M3,3H21V21H3V3M7,7V17H17V7H7M7,9H17V11H7V9M7,13H17V15H7V13Z",
        Color = "#F43F5E",
        Description = "Inserts a new row into the table (above by default).",
        DetailedHelpText = "### Insert Table Row\n\nSelect one or more cells and run the command to insert a row above the topmost selected row. Right-click to insert above or below, preserve the total table height, or duplicate the adjacent row's basic formatting into an empty row. Each successful action creates one PowerPoint Undo entry.",
        MinSelection = 0,
        RequiresTable = true)]
    public static class TableRowInsertionFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - inserts row above without preserving height.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, top: true, preserveHeight: false);
        }

        /// <summary>
        /// Inserts a new table row.
        /// </summary>
        /// <param name="manager">The PowerPointManager instance.</param>
        /// <param name="top">True to insert above, False to insert below.</param>
        /// <param name="preserveHeight">True to keep total table height same, False to let it grow.</param>
        /// <returns>True if row was inserted.</returns>
        public static bool Execute(PowerPointManager manager, bool top, bool preserveHeight)
        {
            return ExecuteMultiple(manager, top, preserveHeight, 1);
        }

        public static bool ExecuteMultiple(PowerPointManager manager, bool top, bool preserveHeight, int count)
        {
            if (manager == null) return false;
            if (count < 1 || count > 50)
            {
                MessageBox.Show("Enter a number of rows between 1 and 50.", "Insert Table Rows", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            try
            {
                Shape tableShape;
                Table table;
                if (!TableHelper.GetSelectedTable(manager, out tableShape, out table)) return false;

                int minRow, maxRow, minCol, maxCol;
                if (!TableHelper.GetSelectedCellBounds(table, out minRow, out maxRow, out minCol, out maxCol))
                {
                    MessageBox.Show("Select at least one table cell before inserting a row.", "Insert Table Row", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }

                int targetIndex = top ? minRow : (maxRow + 1);
                float originalHeight = tableShape.Height;
                float originalTop = tableShape.Top;
                float[] originalRowHeights = ReadRowHeights(table);

                int insertedIndex = targetIndex > table.Rows.Count ? table.Rows.Count + 1 : targetIndex;
                int insertedCount = 0;
                try
                {
                    try { manager.GetApplication().StartNewUndoEntry(); }
                    catch (Exception undoEx) { ExceptionPolicy.ReportOptionalComFailure(undoEx, "TableRowInsertionFeature.Insert.StartUndoEntry"); }
                    for (int index = 0; index < count; index++)
                    {
                        if (targetIndex > table.Rows.Count) table.Rows.Add(-1);
                        else table.Rows.Add(targetIndex);
                        insertedCount++;
                    }
                    if (preserveHeight)
                    {
                        ApplyRowHeights(
                            table,
                            TableDimensionScaler.ScaleToTotal(
                                ReadRowHeights(table),
                                originalHeight));
                        tableShape.Top = originalTop;
                        EnsureTotalHeight(table, originalHeight);
                    }
                    return true;
                }
                catch (Exception ex)
                {
                    while (insertedCount-- > 0)
                    {
                        try { table.Rows[insertedIndex].Delete(); } catch (Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "TableRowInsertionFeature.Rollback"); }
                    }
                    try
                    {
                        ApplyRowHeights(table, originalRowHeights);
                        tableShape.Top = originalTop;
                    }
                    catch (Exception rollbackEx) { ExceptionPolicy.ReportOptionalComFailure(rollbackEx, "TableRowInsertionFeature.Insert.RestoreHeight"); }
                    ExceptionLogger.Log(ex, "TableRowInsertionFeature.Insert");
                    MessageBox.Show("The rows could not be inserted. The table was restored where PowerPoint allowed it.", "Insert Table Rows", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableRowInsertionFeature.Execute");
                return false;
            }
        }

        private static float[] ReadRowHeights(Table table)
        {
            var heights = new float[table.Rows.Count];
            for (int index = 1; index <= table.Rows.Count; index++)
                heights[index - 1] = table.Rows[index].Height;
            return heights;
        }

        private static void ApplyRowHeights(Table table, IReadOnlyList<float> heights)
        {
            if (table.Rows.Count != heights.Count)
                throw new InvalidOperationException("The saved row heights no longer match the table.");
            for (int index = 1; index <= table.Rows.Count; index++)
                table.Rows[index].Height = heights[index - 1];
        }

        private static void EnsureTotalHeight(Table table, float expectedHeight)
        {
            float actualHeight = 0f;
            for (int index = 1; index <= table.Rows.Count; index++)
                actualHeight += table.Rows[index].Height;
            if (Math.Abs(actualHeight - expectedHeight) > 0.5f)
                throw new InvalidOperationException("PowerPoint could not preserve the total table height.");
        }

        public static bool DuplicateRowFormat(PowerPointManager manager, bool top)
        {
            if (manager == null) return false;
            try
            {
                Shape tableShape;
                Table table;
                if (!TableHelper.GetSelectedTable(manager, out tableShape, out table)) return false;

                int minRow, maxRow, minCol, maxCol;
                if (!TableHelper.GetSelectedCellBounds(table, out minRow, out maxRow, out minCol, out maxCol))
                {
                    MessageBox.Show("Select at least one table cell before duplicating a row format.", "Duplicate Row Format", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }

                int sourceIndex = top ? minRow : maxRow;
                int targetIndex = top ? minRow : maxRow + 1;
                int insertedIndex = targetIndex > table.Rows.Count ? table.Rows.Count + 1 : targetIndex;
                float sourceHeight = table.Rows[sourceIndex].Height;
                float originalHeight = tableShape.Height;
                bool inserted = false;
                try
                {
                    try { manager.GetApplication().StartNewUndoEntry(); }
                    catch (Exception undoEx) { ExceptionPolicy.ReportOptionalComFailure(undoEx, "TableRowInsertionFeature.Duplicate.StartUndoEntry"); }
                    if (targetIndex > table.Rows.Count) table.Rows.Add(-1);
                    else table.Rows.Add(targetIndex);
                    inserted = true;
                    table.Rows[insertedIndex].Height = sourceHeight;
                    int copySourceIndex = top ? sourceIndex + 1 : sourceIndex;
                    for (int column = 1; column <= table.Columns.Count; column++)
                        TableColumnInsertionFeature.CopyCellFormat(table.Cell(copySourceIndex, column), table.Cell(insertedIndex, column));
                    return true;
                }
                catch (Exception ex)
                {
                    if (inserted)
                    {
                        try { table.Rows[insertedIndex].Delete(); } catch (Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "TableRowInsertionFeature.DuplicateRollback"); }
                        try { tableShape.Height = originalHeight; }
                        catch (Exception sizeEx) { ExceptionPolicy.ReportOptionalComFailure(sizeEx, "TableRowInsertionFeature.Duplicate.RestoreHeight"); }
                    }
                    ExceptionLogger.Log(ex, "TableRowInsertionFeature.DuplicateRowFormat");
                    MessageBox.Show("The row format could not be duplicated. The table was restored where PowerPoint allowed it.", "Duplicate Row Format", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableRowInsertionFeature.DuplicateRowFormat");
                return false;
            }
        }
    }
}
