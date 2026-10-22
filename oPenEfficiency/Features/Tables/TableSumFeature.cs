using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnTableSum",
        Name = "Calculate Table Totals",
        SearchAliases = new[] { "Table Sum" },
        Tooltip = "Calculate table row and column totals",
        IconData = "M3,3H21V21H3V3M7,7H19V19H7V7M9,9V17H17V9H9M11,11H15V13H11V11Z",
        Color = "#F43F5E",
        Description = "Writes row and column totals into the last selected column and row while validating all input values first.",
        DetailedHelpText = "### Calculate Table Totals\n\n**How to use:** Click into a table and select the cell range that contains your numbers plus the row and/or column where the totals belong — for example the last column, the last row, or both. Then run Calculate Table Totals. Row totals are written into the last selected column, column totals into the last selected row, and the bottom-right cell receives the grand total. To calculate column totals only, use the context action instead of the primary command.\n\nThe selection needs at least two rows for column totals and at least two columns for row totals. Empty input cells are ignored; a non-empty value that cannot be parsed in the current, invariant, English, or German number format aborts the complete operation instead of being treated as zero. Outputs are prepared before any cell changes, destination formatting stays in place, and successful writes create a PowerPoint Undo entry.",
        MinSelection = 0,
        RequiresTable = true)]
    public static class TableSumFeature
    {
        public static bool Execute(PowerPointManager manager) => Execute(manager, sumRow: true, sumCol: true);

        public static bool Execute(PowerPointManager manager, bool sumRow, bool sumCol)
        {
            if (manager == null) return false;
            try
            {
                Shape tableShape;
                Table table;
                if (!TableHelper.GetSelectedTable(manager, out tableShape, out table)) return false;

                int rMin, rMax, cMin, cMax;
                if (!TableHelper.GetSelectedCellBounds(table, out rMin, out rMax, out cMin, out cMax))
                {
                    ShowInfo("Select a table range that includes a final row or column for totals.");
                    return false;
                }

                bool calculateColumns = sumCol && rMax > rMin;
                bool calculateRows = sumRow && cMax > cMin;
                if (!calculateColumns && !calculateRows)
                {
                    ShowInfo("The selection needs at least two rows for column totals or two columns for row totals.");
                    return false;
                }

                var cells = new string[rMax - rMin + 1, cMax - cMin + 1];
                for (int row = rMin; row <= rMax; row++)
                {
                    for (int column = cMin; column <= cMax; column++)
                    {
                        cells[row - rMin, column - cMin] =
                            table.Cell(row, column).Shape.TextFrame.TextRange.Text;
                    }
                }
                var outputs = TableTextPlanner
                    .CalculateTotals(cells, calculateRows, calculateColumns, rMin, cMin)
                    .ToDictionary(
                        total => Key(rMin + total.Row, cMin + total.Column),
                        total => total.Value);

                var original = new Dictionary<string, string>();
                foreach (var output in outputs)
                {
                    ParseKey(output.Key, out int row, out int column);
                    original[output.Key] = table.Cell(row, column).Shape.TextFrame.TextRange.Text;
                }

                try { manager.GetApplication().StartNewUndoEntry(); } catch { /* Formula insertion remains valid when undo grouping is unavailable. */ }
                try
                {
                    foreach (var output in outputs)
                    {
                        ParseKey(output.Key, out int row, out int column);
                        table.Cell(row, column).Shape.TextFrame.TextRange.Text = output.Value.ToString("G15", CultureInfo.CurrentCulture);
                    }
                    return true;
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "TableSumFeature.Write");
                    Restore(table, original);
                    System.Windows.Forms.MessageBox.Show("The totals could not be written completely. The destination texts were restored where PowerPoint allowed it.", "Calculate Table Totals", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    return false;
                }
            }
            catch (FormatException ex)
            {
                ExceptionLogger.Log(ex, "TableSumFeature.InvalidInput");
                System.Windows.Forms.MessageBox.Show(ex.Message, "Calculate Table Totals", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableSumFeature.Execute");
                return false;
            }
        }

        private static void Restore(Table table, Dictionary<string, string> original)
        {
            foreach (var item in original)
            {
                ParseKey(item.Key, out int row, out int column);
                try { table.Cell(row, column).Shape.TextFrame.TextRange.Text = item.Value; }
                catch (Exception ex) { ExceptionLogger.Log(ex, "TableSumFeature.Rollback"); }
            }
        }

        private static string Key(int row, int column) => row + ":" + column;

        private static void ParseKey(string key, out int row, out int column)
        {
            var parts = key.Split(':');
            row = int.Parse(parts[0], CultureInfo.InvariantCulture);
            column = int.Parse(parts[1], CultureInfo.InvariantCulture);
        }

        private static void ShowInfo(string message)
        {
            System.Windows.Forms.MessageBox.Show(message, "Calculate Table Totals", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        }
    }
}
