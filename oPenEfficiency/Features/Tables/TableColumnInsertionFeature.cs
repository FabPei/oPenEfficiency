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
    /// Feature: Insert a new column into a table relative to the selection.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnTableColumnInsertion",
        Name = "Insert Table Column",
        SearchAliases = new[] { "Insert Column" },
        Tooltip = "Insert a table column at the selected position",
        IconData = "M3,3H21V21H3V3M9,5V19H11V5H9M13,5V19H15V5H13M7,7V17H17V7H7Z",
        Color = "#F43F5E",
        Description = "Inserts a new column into the table (left by default).",
        DetailedHelpText = "### Insert Table Column\n\nSelect one or more cells and run the command to insert a column before the leftmost selected column. Right-click to insert left or right, preserve the total table width, or duplicate the adjacent column's basic formatting into an empty column. Each successful action creates one PowerPoint Undo entry.",
        MinSelection = 0,
        RequiresTable = true)]
    public static class TableColumnInsertionFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - inserts column to the left without preserving width.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, left: true, preserveWidth: false);
        }

        /// <summary>
        /// Inserts a new table column.
        /// </summary>
        /// <param name="manager">The PowerPointManager instance.</param>
        /// <param name="left">True to insert to the left, False to insert to the right.</param>
        /// <param name="preserveWidth">True to keep total table width same, False to let it grow.</param>
        /// <returns>True if column was inserted.</returns>
        public static bool Execute(PowerPointManager manager, bool left, bool preserveWidth)
        {
            return ExecuteMultiple(manager, left, preserveWidth, 1);
        }

        public static bool ExecuteMultiple(PowerPointManager manager, bool left, bool preserveWidth, int count)
        {
            if (manager == null) return false;
            if (count < 1 || count > 50)
            {
                MessageBox.Show("Enter a number of columns between 1 and 50.", "Insert Table Columns", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            try
            {
                Shape tableShape;
                Table table;
                if (!TableHelper.GetSelectedTable(manager, out tableShape, out table)) return false;

                int minCol, maxCol, minRow, maxRow;
                if (!TableHelper.GetSelectedCellBounds(table, out minRow, out maxRow, out minCol, out maxCol))
                {
                    MessageBox.Show("Select at least one table cell before inserting a column.", "Insert Table Column", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }

                int targetIndex = left ? minCol : (maxCol + 1);
                float originalWidth = tableShape.Width;
                float originalLeft = tableShape.Left;
                float[] originalColumnWidths = ReadColumnWidths(table);

                int insertedIndex = targetIndex > table.Columns.Count ? table.Columns.Count + 1 : targetIndex;
                int insertedCount = 0;
                try
                {
                    try { manager.GetApplication().StartNewUndoEntry(); }
                    catch (Exception undoEx) { ExceptionPolicy.ReportOptionalComFailure(undoEx, "TableColumnInsertionFeature.Insert.StartUndoEntry"); }
                    for (int index = 0; index < count; index++)
                    {
                        if (targetIndex > table.Columns.Count) table.Columns.Add(-1);
                        else table.Columns.Add(targetIndex);
                        insertedCount++;
                    }
                    if (preserveWidth)
                    {
                        ApplyColumnWidths(
                            table,
                            TableDimensionScaler.ScaleToTotal(
                                ReadColumnWidths(table),
                                originalWidth));
                        tableShape.Left = originalLeft;
                        EnsureTotalWidth(table, originalWidth);
                    }
                    return true;
                }
                catch (Exception ex)
                {
                    while (insertedCount-- > 0)
                    {
                        try { table.Columns[insertedIndex].Delete(); } catch (Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "TableColumnInsertionFeature.Rollback"); }
                    }
                    try
                    {
                        ApplyColumnWidths(table, originalColumnWidths);
                        tableShape.Left = originalLeft;
                    }
                    catch (Exception rollbackEx) { ExceptionPolicy.ReportOptionalComFailure(rollbackEx, "TableColumnInsertionFeature.Insert.RestoreWidth"); }
                    ExceptionLogger.Log(ex, "TableColumnInsertionFeature.Insert");
                    MessageBox.Show("The columns could not be inserted. The table was restored where PowerPoint allowed it.", "Insert Table Columns", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableColumnInsertionFeature.Execute");
                return false;
            }
        }

        private static float[] ReadColumnWidths(Table table)
        {
            var widths = new float[table.Columns.Count];
            for (int index = 1; index <= table.Columns.Count; index++)
                widths[index - 1] = table.Columns[index].Width;
            return widths;
        }

        private static void ApplyColumnWidths(Table table, IReadOnlyList<float> widths)
        {
            if (table.Columns.Count != widths.Count)
                throw new InvalidOperationException("The saved column widths no longer match the table.");
            for (int index = 1; index <= table.Columns.Count; index++)
                table.Columns[index].Width = widths[index - 1];
        }

        private static void EnsureTotalWidth(Table table, float expectedWidth)
        {
            float actualWidth = 0f;
            for (int index = 1; index <= table.Columns.Count; index++)
                actualWidth += table.Columns[index].Width;
            if (Math.Abs(actualWidth - expectedWidth) > 0.5f)
                throw new InvalidOperationException("PowerPoint could not preserve the total table width.");
        }

        public static bool DuplicateColumnFormat(PowerPointManager manager, bool left)
        {
            if (manager == null) return false;

            try
            {
                Shape tableShape;
                Table table;
                if (!TableHelper.GetSelectedTable(manager, out tableShape, out table)) return false;

                int minCol, maxCol, minRow, maxRow;
                if (!TableHelper.GetSelectedCellBounds(table, out minRow, out maxRow, out minCol, out maxCol))
                {
                    MessageBox.Show("Select at least one table cell before duplicating a column format.", "Duplicate Column Format", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return false;
                }

                int sourceIndex = left ? minCol : maxCol;
                int targetIndex = left ? minCol : maxCol + 1;
                int insertedIndex = targetIndex > table.Columns.Count ? table.Columns.Count + 1 : targetIndex;
                float sourceWidth = table.Columns[sourceIndex].Width;
                float originalWidth = tableShape.Width;
                bool inserted = false;

                try
                {
                    try { manager.GetApplication().StartNewUndoEntry(); }
                    catch (Exception undoEx) { ExceptionPolicy.ReportOptionalComFailure(undoEx, "TableColumnInsertionFeature.Duplicate.StartUndoEntry"); }
                    if (targetIndex > table.Columns.Count) table.Columns.Add(-1);
                    else table.Columns.Add(targetIndex);
                    inserted = true;

                    table.Columns[insertedIndex].Width = sourceWidth;
                    int copySourceIndex = left ? sourceIndex + 1 : sourceIndex;
                    for (int row = 1; row <= table.Rows.Count; row++)
                        CopyCellFormat(table.Cell(row, copySourceIndex), table.Cell(row, insertedIndex));
                    return true;
                }
                catch (Exception ex)
                {
                    if (inserted)
                    {
                        try { table.Columns[insertedIndex].Delete(); } catch (Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "TableColumnInsertionFeature.DuplicateRollback"); }
                        try { tableShape.Width = originalWidth; }
                        catch (Exception sizeEx) { ExceptionPolicy.ReportOptionalComFailure(sizeEx, "TableColumnInsertionFeature.Duplicate.RestoreWidth"); }
                    }
                    ExceptionLogger.Log(ex, "TableColumnInsertionFeature.DuplicateColumnFormat");
                    MessageBox.Show("The column format could not be duplicated. The table was restored where PowerPoint allowed it.", "Duplicate Column Format", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableColumnInsertionFeature.DuplicateColumnFormat");
                return false;
            }
        }

        internal static void CopyCellFormat(Cell source, Cell target)
        {
            var sourceShape = source.Shape;
            var targetShape = target.Shape;

            // PowerPoint's Shape.PickUp()/Apply() throws NotImplementedException on table cells
            // (HRESULT 0x80004001), so the paint-pair pickup cannot be used here. Copy the fill
            // explicitly with the same proven idiom TableFormatPainterFeature uses; fonts,
            // text-frame properties and borders are copied below anyway.
            targetShape.Fill.Solid();
            targetShape.Fill.Visible = sourceShape.Fill.Visible;
            targetShape.Fill.ForeColor.RGB = sourceShape.Fill.ForeColor.RGB;
            targetShape.Fill.Transparency = sourceShape.Fill.Transparency;
            targetShape.TextFrame.TextRange.Text = string.Empty;

            var sourceFont = sourceShape.TextFrame.TextRange.Font;
            var targetFont = targetShape.TextFrame.TextRange.Font;
            targetFont.Name = sourceFont.Name;
            targetFont.Size = sourceFont.Size;
            targetFont.Color.RGB = sourceFont.Color.RGB;
            targetFont.Bold = sourceFont.Bold;
            targetFont.Italic = sourceFont.Italic;
            targetFont.Underline = sourceFont.Underline;
            targetShape.TextFrame.TextRange.ParagraphFormat.Alignment = sourceShape.TextFrame.TextRange.ParagraphFormat.Alignment;
            targetShape.TextFrame.VerticalAnchor = sourceShape.TextFrame.VerticalAnchor;
            targetShape.TextFrame.Orientation = sourceShape.TextFrame.Orientation;
            targetShape.TextFrame.MarginLeft = sourceShape.TextFrame.MarginLeft;
            targetShape.TextFrame.MarginRight = sourceShape.TextFrame.MarginRight;
            targetShape.TextFrame.MarginTop = sourceShape.TextFrame.MarginTop;
            targetShape.TextFrame.MarginBottom = sourceShape.TextFrame.MarginBottom;

            CopyBorder(source, target, PpBorderType.ppBorderTop);
            CopyBorder(source, target, PpBorderType.ppBorderBottom);
            CopyBorder(source, target, PpBorderType.ppBorderLeft);
            CopyBorder(source, target, PpBorderType.ppBorderRight);
        }

        private static void CopyBorder(Cell source, Cell target, PpBorderType borderType)
        {
            var sourceBorder = source.Borders[borderType];
            var targetBorder = target.Borders[borderType];
            targetBorder.Visible = sourceBorder.Visible;
            targetBorder.ForeColor.RGB = sourceBorder.ForeColor.RGB;
            targetBorder.Weight = sourceBorder.Weight;
            targetBorder.Style = sourceBorder.Style;
            targetBorder.DashStyle = sourceBorder.DashStyle;
        }
    }
}
