using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Helper class for table row data during sorting.
    /// </summary>
    public class TableRowData
    {
        public string[] Cells { get; set; }
        public string SortValue { get; set; }
    }

    /// <summary>
    /// Feature: Sort table cells or whole rows based on selection.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnTableSortAZ",
        Name = "Sort A-Z",
        Tooltip = "Sort A-Z",
        IconData = "M3,3H21V21H3V3M7,5V19L11,15H8V13H13V16L17,12V5H7M15,7H17V11H19V7H21V5H15V7Z",
        Color = "#F43F5E",
        Description = "Sorts table rows in ascending order (A-Z) based on selected cell column.",
        DetailedHelpText = "### Sort A-Z\nSorts the rows of the selected table alphabetically or numerically based on a chosen column, with ascending and descending support.",
        Keywords = "sort, order, alphabet, arrange, list, organize",
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

        /// <summary>
        /// Wrapper for auto-discovery - sorts table A-Z based on selection.
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, SortType.AZ_Selection);
        }

        /// <summary>
        /// Executes table sorting based on the specified sort type.
        /// </summary>
        public static bool Execute(PowerPointManager manager, SortType type)
        {
            if (manager == null) return false;

            bool ascending = (type == SortType.AZ_Selection || type == SortType.AZ_WholeRow);
            bool wholeRow = (type == SortType.AZ_WholeRow || type == SortType.ZA_WholeRow);

            try
            {
                Shape tableShape;
                Table table;
                if (!TableHelper.GetSelectedTable(manager, out tableShape, out table)) return false;

                int minRow, maxRow, minCol, maxCol;
                if (!TableHelper.GetSelectedCellBounds(table, out minRow, out maxRow, out minCol, out maxCol)) return false;

                int startRow = minRow;
                int endRow = maxRow;
                int sortColumn = minCol;

                if (startRow >= endRow) return true;

                int colCount = table.Columns.Count;
                var rowDataList = new List<TableRowData>();

                for (int r = startRow; r <= endRow; r++)
                {
                    var data = new string[colCount];
                    if (wholeRow)
                    {
                        for (int c = 1; c <= colCount; c++)
                        {
                            try { data[c - 1] = table.Cell(r, c).Shape.TextFrame.TextRange.Text; }
                            catch (Exception ex)
                            {
                                ExceptionLogger.Log(ex, $"TableSort: Error reading cell [{r},{c}]");
                                data[c - 1] = "";
                            }
                        }
                    }
                    else
                    {
                        try { data[sortColumn - 1] = table.Cell(r, sortColumn).Shape.TextFrame.TextRange.Text; }
                        catch (Exception ex)
                        {
                            ExceptionLogger.Log(ex, $"TableSort: Error reading sort cell [{r},{sortColumn}]");
                            data[sortColumn - 1] = "";
                        }
                    }

                    rowDataList.Add(new TableRowData { Cells = data, SortValue = data[sortColumn - 1] });
                }

                rowDataList.Sort((a, b) =>
                {
                    int res = string.Compare(a.SortValue, b.SortValue, StringComparison.CurrentCultureIgnoreCase);
                    return ascending ? res : -res;
                });

                for (int i = 0; i < rowDataList.Count; i++)
                {
                    int r = startRow + i;
                    if (wholeRow)
                    {
                        for (int c = 1; c <= colCount; c++)
                        {
                            try { table.Cell(r, c).Shape.TextFrame.TextRange.Text = rowDataList[i].Cells[c - 1]; }
                            catch (Exception ex)
                            {
                                ExceptionLogger.Log(ex, $"TableSort: Error writing cell [{r},{c}]");
                            }
                        }
                    }
                    else
                    {
                        try { table.Cell(r, sortColumn).Shape.TextFrame.TextRange.Text = rowDataList[i].Cells[sortColumn - 1]; }
                        catch (Exception ex)
                        {
                            ExceptionLogger.Log(ex, $"TableSort: Error writing cell [{r},{sortColumn}]");
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableSortFeature.Execute");
                return false;
            }
        }
    }
}
