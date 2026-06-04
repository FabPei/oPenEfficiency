using System;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Split a table into two pieces based on the current selection.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnTableSplit",
        Name = "Split Table",
        Tooltip = "Split table",
        IconData = "M3,3H21V21H3V3M7,5V19H10V5H7M12,5V19H14V5H12M15,5V19H19V5H15Z",
        Color = "#F43F5E",
        Description = "Splits the table into two separate tables at the selected row or column.",
        DetailedHelpText = "### Split Table\nSplits the selected table into two separate tables at the currently selected row. Upper and lower halves are repositioned automatically.",
        Keywords = "split, divide, separate, cut, break, table",
        MinSelection = 0,
        RequiresTable = true)]
    public static class TableSplitFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - splits table by column.
        /// </summary>
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

                int splitIndex = -1;
                for (int r = 1; r <= table.Rows.Count; r++)
                {
                    for (int c = 1; c <= table.Columns.Count; c++)
                    {
                        if (table.Cell(r, c).Selected)
                        {
                            splitIndex = byColumn ? c : r;
                            break;
                        }
                    }
                    if (splitIndex != -1) break;
                }

                if (splitIndex <= 1 || (byColumn && splitIndex > table.Columns.Count) || (!byColumn && splitIndex > table.Rows.Count))
                    return false;

                var dupRange = tableShape.Duplicate();
                var table2Shape = dupRange[1];
                var table2 = table2Shape.Table;

                if (byColumn)
                {
                    for (int c = table.Columns.Count; c >= splitIndex; c--) table.Columns[c].Delete();
                    for (int c = splitIndex - 1; c >= 1; c--) table2.Columns[c].Delete();

                    table2Shape.Left = tableShape.Left + tableShape.Width + 20;
                    table2Shape.Top = tableShape.Top;
                }
                else
                {
                    for (int r = table.Rows.Count; r >= splitIndex; r--) table.Rows[r].Delete();
                    for (int r = splitIndex - 1; r >= 1; r--) table2.Rows[r].Delete();

                    table2Shape.Left = tableShape.Left;
                    table2Shape.Top = tableShape.Top + tableShape.Height + 20;
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableSplitFeature.Execute");
                return false;
            }
        }
    }
}
