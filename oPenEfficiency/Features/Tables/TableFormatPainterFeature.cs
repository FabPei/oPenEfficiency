using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Propagate formatting (fill, text style) from selected table cells to neighbors.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnTableFormatPainter",
        Name = "Table Format Painter",
        Tooltip = "Table Format Painter",
        Color = "#06B6D4",
        Description = "Quickly paint background and text formatting from selected table cells to neighboring cells in specific directions (Up, Down, Left, Right).",
        DetailedHelpText = "### Table Format Painter\nCopies the fill color, font, and text alignment from one selected table cell and applies it to all other cells in a chosen direction.",
        MinSelection = 0,
        RequiresTable = true)]
    public static class TableFormatPainterFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            var win = new UI.TableFormatPainterWindow(manager);
            win.Show();
            return true;
        }

        /// <summary>
        /// Performs the actual paint logic.
        /// </summary>
        public static bool Paint(PowerPointManager manager, bool up, bool down, bool left, bool right, bool formatText, bool formatCell, int skip)
        {
            if (manager == null) return false;

            try
            {
                Shape tableShape;
                Table table;
                if (!TableHelper.GetSelectedTable(manager, out tableShape, out table)) return false;

                var selectedCells = new List<(int R, int C)>();
                for (int r = 1; r <= table.Rows.Count; r++)
                {
                    for (int c = 1; c <= table.Columns.Count; c++)
                    {
                        if (table.Cell(r, c).Selected)
                            selectedCells.Add((r, c));
                    }
                }

                if (selectedCells.Count == 0) return false;

                foreach (var sc in selectedCells)
                {
                    if (up) PaintDirection(table, sc.R, sc.C, -1, 0, formatText, formatCell, skip);
                    if (down) PaintDirection(table, sc.R, sc.C, 1, 0, formatText, formatCell, skip);
                    if (left) PaintDirection(table, sc.R, sc.C, 0, -1, formatText, formatCell, skip);
                    if (right) PaintDirection(table, sc.R, sc.C, 0, 1, formatText, formatCell, skip);
                }

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableFormatPainterFeature.Paint");
                return false;
            }
        }

        private static void PaintDirection(Table table, int startR, int startC, int dr, int dc, bool text, bool cell, int skip)
        {
            var source = table.Cell(startR, startC);
            int currR = startR + dr;
            int currC = startC + dc;
            int skipCount = 0;

            while (currR >= 1 && currR <= table.Rows.Count && currC >= 1 && currC <= table.Columns.Count)
            {
                if (skipCount == 0)
                {
                    var target = table.Cell(currR, currC);
                    if (cell)
                    {
                        target.Shape.Fill.ForeColor.RGB = source.Shape.Fill.ForeColor.RGB;
                        target.Shape.Fill.Transparency = source.Shape.Fill.Transparency;
                    }
                    if (text)
                    {
                        var sTR = source.Shape.TextFrame.TextRange.Font;
                        var tTR = target.Shape.TextFrame.TextRange.Font;
                        tTR.Name = sTR.Name;
                        tTR.Size = sTR.Size;
                        tTR.Color.RGB = sTR.Color.RGB;
                        tTR.Bold = sTR.Bold;
                        tTR.Italic = sTR.Italic;
                    }
                    skipCount = skip;
                }
                else
                {
                    skipCount--;
                }

                currR += dr;
                currC += dc;
            }
        }
    }
}
