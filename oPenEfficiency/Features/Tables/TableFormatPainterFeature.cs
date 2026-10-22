using System;
using System.Collections.Generic;
using System.Windows.Interop;
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
        Tooltip = "Copy cell fill or text formatting along rows and columns",
        Color = "#06B6D4",
        Description = "Paints fill or basic text formatting from one selected table cell along chosen row and column directions.",
        DetailedHelpText = "### Table Format Painter\n\nSelect exactly one source cell, then choose one or more directions. Right-click for zero-skip horizontal, vertical, or all-direction presets that copy text formatting, cell fill, or both. **Text Formatting** copies the basic font, paragraph alignment, vertical anchor, and margins; **Cell Fill** copies a visible solid fill and transparency. Borders, rich-text runs, hyperlinks, gradients, patterns, and effects are not copied. **Cells to skip between paints** controls the interval: 0 paints every cell, 1 every second cell, and so on. The palette is bound to PowerPoint and only one instance opens at a time.",
        MinSelection = 0,
        RequiresTable = true)]
    public static class TableFormatPainterFeature
    {
        private static UI.TableFormatPainterWindow _window;

        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                if (_window != null && _window.IsLoaded)
                {
                    _window.Activate();
                    return true;
                }

                _window = new UI.TableFormatPainterWindow(manager);
                var app = manager.GetApplication();
                if (app != null) new WindowInteropHelper(_window).Owner = (IntPtr)app.HWND;
                _window.Closed += (s, e) => _window = null;
                _window.Show();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableFormatPainterFeature.Execute");
                var failedWindow = _window;
                _window = null;
                try { failedWindow?.Close(); }
                catch (Exception closeEx)
                {
                    ExceptionLogger.Log(closeEx, "TableFormatPainterFeature.Execute.FailureCleanup");
                }
                throw;
            }
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

                if (selectedCells.Count != 1)
                {
                    System.Windows.Forms.MessageBox.Show("Select exactly one source cell before painting table formats.", "Table Format Painter", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return false;
                }

                if (!up && !down && !left && !right) return false;
                if (!formatText && !formatCell) return false;
                if (skip < 0) return false;

                try { manager.GetApplication().StartNewUndoEntry(); } catch { /* Painting remains valid when undo grouping is unavailable. */ }
                var painted = new HashSet<string>();

                foreach (var sc in selectedCells)
                {
                    if (up) PaintDirection(table, sc.R, sc.C, -1, 0, formatText, formatCell, skip, painted);
                    if (down) PaintDirection(table, sc.R, sc.C, 1, 0, formatText, formatCell, skip, painted);
                    if (left) PaintDirection(table, sc.R, sc.C, 0, -1, formatText, formatCell, skip, painted);
                    if (right) PaintDirection(table, sc.R, sc.C, 0, 1, formatText, formatCell, skip, painted);
                }

                if (painted.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show("No target cells exist in the selected direction and interval.", "Table Format Painter", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableFormatPainterFeature.Paint");
                System.Windows.Forms.MessageBox.Show("The formats could not be painted completely. Use PowerPoint Undo to restore the previous formatting.", "Table Format Painter", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }

        private static void PaintDirection(Table table, int startR, int startC, int dr, int dc, bool text, bool cell, int skip, HashSet<string> painted)
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
                    string key = currR + ":" + currC;
                    if (!painted.Add(key))
                    {
                        currR += dr;
                        currC += dc;
                        continue;
                    }
                    if (cell)
                    {
                        target.Shape.Fill.Solid();
                        target.Shape.Fill.Visible = source.Shape.Fill.Visible;
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
                        tTR.Underline = sTR.Underline;
                        target.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = source.Shape.TextFrame.TextRange.ParagraphFormat.Alignment;
                        target.Shape.TextFrame.VerticalAnchor = source.Shape.TextFrame.VerticalAnchor;
                        target.Shape.TextFrame.Orientation = source.Shape.TextFrame.Orientation;
                        target.Shape.TextFrame.MarginLeft = source.Shape.TextFrame.MarginLeft;
                        target.Shape.TextFrame.MarginRight = source.Shape.TextFrame.MarginRight;
                        target.Shape.TextFrame.MarginTop = source.Shape.TextFrame.MarginTop;
                        target.Shape.TextFrame.MarginBottom = source.Shape.TextFrame.MarginBottom;
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
