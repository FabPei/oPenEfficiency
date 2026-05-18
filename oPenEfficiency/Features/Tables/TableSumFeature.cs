using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Calculate the sum of selected table cells and place the result in the last selected row/column.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnTableSum",
        Name = "Table Sum",
        Tooltip = "Table sum",
        IconData = "M3,3H21V21H3V3M7,7H19V19H7V7M9,9V17H17V9H9M11,11H15V13H11V11Z",
        Color = "#F43F5E",
        Description = "Calculates the sum of selected table cells.",
        DetailedHelpText = "### Table Sum\nInserts the calculated sum of all numeric values in the selected table column or row into a configurable target cell.",
        MinSelection = 0,
        RequiresTable = true)]
    public static class TableSumFeature
    {
        /// <summary>
        /// Wrapper for auto-discovery - sums selected cells (row and column).
        /// </summary>
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, sumRow: true, sumCol: true);
        }

        public static bool Execute(PowerPointManager manager, bool sumRow, bool sumCol)
        {
            if (manager == null) return false;

            try
            {
                Shape tableShape;
                Table table;
                if (!TableHelper.GetSelectedTable(manager, out tableShape, out table)) return false;

                int rMin, rMax, cMin, cMax;
                if (!TableHelper.GetSelectedCellBounds(table, out rMin, out rMax, out cMin, out cMax)) return false;

                if (cMax == cMin && rMax == rMin) return false;

                if (sumCol && rMax > rMin)
                {
                    for (int c = cMin; c <= cMax; c++)
                    {
                        double sum = 0;
                        for (int r = rMin; r < rMax; r++) sum += TableHelper.ParseDoubleFromCell(table.Cell(r, c).Shape.TextFrame.TextRange.Text);
                        table.Cell(rMax, c).Shape.TextFrame.TextRange.Text = sum.ToString(System.Globalization.CultureInfo.CurrentCulture);
                    }
                }

                if (sumRow && cMax > cMin)
                {
                    for (int r = rMin; r <= rMax; r++)
                    {
                        double sum = 0;
                        int limit = (sumCol && r == rMax) ? cMax - 1 : cMax - 1;
                        for (int c = cMin; c <= limit; c++) sum += TableHelper.ParseDoubleFromCell(table.Cell(r, c).Shape.TextFrame.TextRange.Text);
                        table.Cell(r, cMax).Shape.TextFrame.TextRange.Text = sum.ToString(System.Globalization.CultureInfo.CurrentCulture);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableSumFeature.Execute");
                return false;
            }
        }
    }
}
