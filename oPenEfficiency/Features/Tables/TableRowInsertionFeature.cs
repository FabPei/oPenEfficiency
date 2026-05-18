using System;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Insert a new row into a table relative to the selection.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnTableRowInsertion",
        Name = "Insert Row",
        Tooltip = "Insert row",
        IconData = "M3,3H21V21H3V3M7,7V17H17V7H7M7,9H17V11H7V9M7,13H17V15H7V13Z",
        Color = "#F43F5E",
        Description = "Inserts a new row into the table (above by default).",
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
            if (manager == null) return false;

            try
            {
                Shape tableShape;
                Table table;
                if (!TableHelper.GetSelectedTable(manager, out tableShape, out table)) return false;

                int minRow, maxRow, minCol, maxCol;
                if (!TableHelper.GetSelectedCellBounds(table, out minRow, out maxRow, out minCol, out maxCol)) return false;

                int targetIndex = top ? minRow : (maxRow + 1);
                float originalHeight = tableShape.Height;

                if (targetIndex > table.Rows.Count)
                    table.Rows.Add(-1);
                else
                    table.Rows.Add(targetIndex);

                if (preserveHeight) tableShape.Height = originalHeight;

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableRowInsertionFeature.Execute");
                return false;
            }
        }
    }
}
