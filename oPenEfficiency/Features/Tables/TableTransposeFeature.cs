using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Transpose a table (rows become columns and vice versa).
    /// </summary>
    [FeatureMetadata(
        Id = "BtnTableTranspose",
        Name = "Transpose Table",
        Tooltip = "Transpose table",
        IconData = "M3,3H21V21H3V3M7,5V19H19V5H7M9,7H17V9H9V7M9,11H17V13H9V11M9,15H14V17H9V15Z",
        Color = "#F43F5E",
        Description = "Transposes the table, swapping rows and columns.",
        DetailedHelpText = "### Transpose Table\nTransposes the selected table, converting rows to columns and columns to rows, while preserving text content and formatting.",
        Keywords = "swap rows, columns, flip table, rotate table, matrix transpose",
        MinSelection = 0,
        RequiresTable = true)]
    public static class TableTransposeFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;

            try
            {
                Shape tableShape;
                Table table;
                if (!TableHelper.GetSelectedTable(manager, out tableShape, out table)) return false;

                int origRows = table.Rows.Count;
                int origCols = table.Columns.Count;
                int size = Math.Max(origRows, origCols);

                // Add rows/cols to make it square
                for (int r = origRows + 1; r <= size; r++) table.Rows.Add(-1);
                for (int c = origCols + 1; c <= size; c++) table.Columns.Add(-1);

                // Swap cells (r, c) with (c, r)
                for (int r = 1; r <= size; r++)
                {
                    for (int c = r + 1; c <= size; c++)
                    {
                        var cell1 = table.Cell(r, c);
                        var cell2 = table.Cell(c, r);

                        string text1 = cell1.Shape.TextFrame.TextRange.Text;
                        string text2 = cell2.Shape.TextFrame.TextRange.Text;
                        cell1.Shape.TextFrame.TextRange.Text = text2;
                        cell2.Shape.TextFrame.TextRange.Text = text1;

                        try
                        {
                            bool f1Vis = cell1.Shape.Fill.Visible == Office.MsoTriState.msoTrue;
                            bool f2Vis = cell2.Shape.Fill.Visible == Office.MsoTriState.msoTrue;
                            int f1Rgb = cell1.Shape.Fill.ForeColor.RGB;
                            int f2Rgb = cell2.Shape.Fill.ForeColor.RGB;

                            cell1.Shape.Fill.Visible = f2Vis ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                            if (f2Vis) cell1.Shape.Fill.ForeColor.RGB = f2Rgb;

                            cell2.Shape.Fill.Visible = f1Vis ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                            if (f1Vis) cell2.Shape.Fill.ForeColor.RGB = f1Rgb;
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"TableTranspose: Error swapping cell formats: {ex.Message}");
                        }
                    }
                }

                // Delete excess original
                for (int r = size; r > origCols; r--) table.Rows[r].Delete();
                for (int c = size; c > origRows; c--) table.Columns[c].Delete();

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableTransposeFeature.Execute");
                return false;
            }
        }
    }
}
