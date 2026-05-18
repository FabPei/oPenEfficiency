using System;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Insert a new column into a table relative to the selection.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnTableColumnInsertion",
        Name = "Insert Column",
        Tooltip = "Insert column",
        IconData = "M3,3H21V21H3V3M9,5V19H11V5H9M13,5V19H15V5H13M7,7V17H17V7H7Z",
        Color = "#F43F5E",
        Description = "Inserts a new column into the table (left by default).",
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
            if (manager == null) return false;

            try
            {
                Shape tableShape;
                Table table;
                if (!TableHelper.GetSelectedTable(manager, out tableShape, out table)) return false;

                int minCol, maxCol, minRow, maxRow;
                if (!TableHelper.GetSelectedCellBounds(table, out minRow, out maxRow, out minCol, out maxCol)) return false;

                int targetIndex = left ? minCol : (maxCol + 1);
                float originalWidth = tableShape.Width;

                if (targetIndex > table.Columns.Count)
                    table.Columns.Add(-1);
                else
                    table.Columns.Add(targetIndex);

                if (preserveWidth) tableShape.Width = originalWidth;

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableColumnInsertionFeature.Execute");
                return false;
            }
        }
    }
}
