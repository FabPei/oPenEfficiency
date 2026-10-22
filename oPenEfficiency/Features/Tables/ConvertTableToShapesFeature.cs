using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Convert a PowerPoint table into a set of grouped rectangular shapes.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnConvertTableToShapes",
        Name = "Convert Table to Shapes",
        Tooltip = "Convert the selected table to shapes",
        IconData = "M4,4H20V20H4V4M6,6V10H10V6H6M12,6V10H16V6H12M18,6V10H18V6M6,12V16H10V12H6M12,12V16H16V12H12M18,12V16H18V12M6,18V18H10V18H6M12,18V18H16V18H12M18,18V18H18V18Z",
        Color = "#F43F5E",
        Description = "Converts one PowerPoint table into a group of editable cell shapes while retaining basic text, fill, line, and alignment formatting.",
        DetailedHelpText = "### Convert Table to Shapes\n\nConverts one selected native PowerPoint table into a group of independently editable cell shapes at the same slide coordinates. The primary action asks whether to delete, keep, or discard the conversion. Right-click can choose Keep Original or Replace Original before conversion; replacement still occurs only after all shapes were created successfully. Rotated tables are rejected because their cells cannot be reproduced safely as unrotated rectangles.\n\nThe conversion retains plain text and basic fill, line, font, alignment, orientation, and margin settings. Rich text, hyperlinks, table-specific borders, merged-cell semantics, gradients, patterns, and other advanced table formatting may not be retained. Review the generated shapes before deleting the original table.",
        MinSelection = 1,
        MaxSelection = 1,
        RequiredType = OfficeSelectionType.Shapes,
        RequiresTable = true)]
    public static class ConvertTableToShapesFeature
    {
        public enum SourceDisposition
        {
            AskAfterConversion,
            KeepOriginal,
            ReplaceOriginal
        }

        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, SourceDisposition.AskAfterConversion);
        }

        /// <param name="interactive">
        /// False selects the message-box-free programmatic path (acceptance seam): the conversion
        /// proceeds without the confirmation prompt, AskAfterConversion skips its follow-up
        /// question (keeping the original unless the disposition says ReplaceOriginal), and no
        /// message box is ever raised. Default is true, so the dialog-driven user path is
        /// unchanged.
        /// </param>
        public static bool Execute(PowerPointManager manager, SourceDisposition disposition, bool interactive = true)
        {
            if (manager == null) return false;

            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;

                Shape tableShape;
                Table table;
                if (!TableHelper.GetSelectedTable(manager, out tableShape, out table)) return false;

                var slide = app.ActiveWindow.View.Slide as Slide;
                if (slide == null) return false;

                if (Math.Abs(tableShape.Rotation) > 0.01f)
                {
                    if (interactive)
                    {
                        System.Windows.Forms.MessageBox.Show("Rotated tables are not supported because PowerPoint exposes their cells as unrotated rectangles. Reset the table rotation to 0° and try again.", "Convert Table to Shapes", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    }
                    return false;
                }

                var cells = PrepareCells(tableShape, table);
                if (cells.Count == 0)
                {
                    if (interactive)
                    {
                        System.Windows.Forms.MessageBox.Show("The selected table does not contain convertible cells.", "Convert Table to Shapes", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    }
                    return false;
                }

                string sourceResult = disposition == SourceDisposition.KeepOriginal
                    ? "The original table will be kept."
                    : disposition == SourceDisposition.ReplaceOriginal
                        ? "The original table will be deleted only after the conversion succeeds."
                        : "After conversion, you can delete the original, keep it, or discard the converted shapes.";
                if (interactive)
                {
                    var confirmation = System.Windows.Forms.MessageBox.Show(
                        "Convert the selected table into editable shapes?\n\n" + sourceResult +
                        "\n\nBasic text and formatting are copied, but rich text, hyperlinks, table-specific borders, merged-cell semantics, gradients, and patterns may not be retained.",
                        "Convert Table to Shapes",
                        System.Windows.Forms.MessageBoxButtons.YesNo,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    if (confirmation != System.Windows.Forms.DialogResult.Yes) return false;
                }

                int targetZ = tableShape.ZOrderPosition;
                string originalName = tableShape.Name;
                var created = new List<Shape>();
                var seen = new HashSet<string>();
                Shape group = null;

                try
                {
                    try { app.StartNewUndoEntry(); } catch { /* Conversion remains valid when undo grouping is unavailable. */ }
                    foreach (var cell in cells)
                    {
                        string key = $"{Math.Round(cell.Left, 2)}|{Math.Round(cell.Top, 2)}|{Math.Round(cell.Width, 2)}|{Math.Round(cell.Height, 2)}";
                        if (seen.Contains(key)) continue;
                        seen.Add(key);

                        var rect = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, cell.Left, cell.Top, cell.Width, cell.Height);
                        rect.Name = $"oPE_Cell_{cell.Row}_{cell.Column}_" + Guid.NewGuid().ToString().Substring(0, 8);
                        CopyBasicAppearance(cell, rect);
                        created.Add(rect);
                    }

                    if (created.Count == 0) throw new InvalidOperationException("No shapes were created from the selected table.");
                    var names = new string[created.Count];
                    for (int index = 0; index < created.Count; index++) names[index] = created[index].Name;
                    group = slide.Shapes.Range(names).Group();
                    group.Name = "oPen_ConvertedTable_" + originalName;

                    int guard = 0;
                    while (group.ZOrderPosition < targetZ && guard < 1000) { group.ZOrder(Office.MsoZOrderCmd.msoBringForward); guard++; }
                    while (group.ZOrderPosition > targetZ && guard < 2000) { group.ZOrder(Office.MsoZOrderCmd.msoSendBackward); guard++; }
                }
                catch (Exception ex)
                {
                    DeleteCreatedShapes(group, created);
                    ExceptionLogger.Log(ex, "ConvertTableToShapesFeature.Convert");
                    if (interactive)
                    {
                        System.Windows.Forms.MessageBox.Show("The table could not be converted. The original table was left unchanged.", "Convert Table to Shapes", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    }
                    return false;
                }

                bool deleteOriginal = disposition == SourceDisposition.ReplaceOriginal;
                if (disposition == SourceDisposition.AskAfterConversion && interactive)
                {
                    var res = System.Windows.Forms.MessageBox.Show(
                        "Conversion succeeded.\n\nYes: delete the original table\nNo: keep the original table visible\nCancel: discard the converted shapes and keep the original table",
                        "Convert Table to Shapes",
                        System.Windows.Forms.MessageBoxButtons.YesNoCancel,
                        System.Windows.Forms.MessageBoxIcon.Question);
                    if (res == System.Windows.Forms.DialogResult.Cancel)
                    {
                        DeleteCreatedShapes(group, created);
                        return false;
                    }
                    deleteOriginal = res == System.Windows.Forms.DialogResult.Yes;
                }
                if (deleteOriginal)
                {
                    try { tableShape.Delete(); }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "ConvertTableToShapesFeature.DeleteOriginal");
                        if (interactive)
                        {
                            System.Windows.Forms.MessageBox.Show("The converted shapes were created, but the original table could not be deleted and remains on the slide.", "Convert Table to Shapes", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                        }
                    }
                }

                try { group.Select(); } catch { /* Result selection is optional after successful conversion. */ }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ConvertTableToShapesFeature.Execute");
                return false;
            }
        }

        private static List<CellConversion> PrepareCells(Shape tableShape, Table table)
        {
            var cells = new List<CellConversion>();
            for (int row = 1; row <= table.Rows.Count; row++)
            {
                for (int column = 1; column <= table.Columns.Count; column++)
                {
                    var sourceCell = table.Cell(row, column);
                    var source = sourceCell.Shape;
                    // Cell.Shape coordinates are already relative to the slide.
                    cells.Add(new CellConversion(sourceCell, source, row, column, source.Left, source.Top, source.Width, source.Height));
                }
            }
            return cells;
        }

        private static void CopyBasicAppearance(CellConversion cell, Shape destination)
        {
            Shape source = cell.Source;
            destination.Fill.Visible = source.Fill.Visible;
            destination.Fill.ForeColor.RGB = source.Fill.ForeColor.RGB;
            destination.Fill.Transparency = source.Fill.Transparency;
            destination.Fill.Solid();

            CopyBasicLineAppearance(cell, destination);

            destination.TextFrame.TextRange.Text = source.TextFrame.TextRange.Text;
            var sourceFont = source.TextFrame.TextRange.Font;
            var destinationFont = destination.TextFrame.TextRange.Font;
            destinationFont.Name = sourceFont.Name;
            destinationFont.Size = sourceFont.Size;
            destinationFont.Color.RGB = sourceFont.Color.RGB;
            destinationFont.Bold = sourceFont.Bold;
            destinationFont.Italic = sourceFont.Italic;
            destination.TextFrame.TextRange.ParagraphFormat.Alignment = source.TextFrame.TextRange.ParagraphFormat.Alignment;
            destination.TextFrame.VerticalAnchor = source.TextFrame.VerticalAnchor;
            destination.TextFrame.Orientation = source.TextFrame.Orientation;
            destination.TextFrame.MarginLeft = source.TextFrame.MarginLeft;
            destination.TextFrame.MarginRight = source.TextFrame.MarginRight;
            destination.TextFrame.MarginTop = source.TextFrame.MarginTop;
            destination.TextFrame.MarginBottom = source.TextFrame.MarginBottom;
        }

        private static void CopyBasicLineAppearance(CellConversion cell, Shape destination)
        {
            LineFormat sourceLine = null;
            try
            {
                sourceLine = cell.Source.Line;
            }
            catch (Exception ex) when (ex is NotImplementedException || ExceptionPolicy.IsExpectedComFailure(ex))
            {
                // Some PowerPoint builds do not expose Shape.Line for table-cell shapes.
            }

            if (sourceLine == null)
            {
                try
                {
                    sourceLine = cell.SourceCell.Borders[PpBorderType.ppBorderTop];
                }
                catch (Exception ex) when (ex is NotImplementedException || ExceptionPolicy.IsExpectedComFailure(ex))
                {
                    // A table-specific border is optional; the converted rectangle remains usable.
                }
            }

            if (sourceLine == null) return;

            try
            {
                destination.Line.Visible = sourceLine.Visible;
                destination.Line.ForeColor.RGB = sourceLine.ForeColor.RGB;
                destination.Line.Weight = sourceLine.Weight;
                destination.Line.DashStyle = sourceLine.DashStyle;
            }
            catch (Exception ex)
            {
                ExceptionPolicy.ReportOptionalComFailure(ex, "ConvertTableToShapesFeature.CopyBasicLineAppearance");
            }
        }

        private static void DeleteCreatedShapes(Shape group, List<Shape> created)
        {
            try
            {
                if (group != null) { group.Delete(); return; }
                foreach (var shape in created) shape.Delete();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ConvertTableToShapesFeature.DeleteCreatedShapes");
            }
        }

        private sealed class CellConversion
        {
            public CellConversion(Cell sourceCell, Shape source, int row, int column, float left, float top, float width, float height)
            {
                SourceCell = sourceCell;
                Source = source;
                Row = row;
                Column = column;
                Left = left;
                Top = top;
                Width = width;
                Height = height;
            }

            public Cell SourceCell { get; }
            public Shape Source { get; }
            public int Row { get; }
            public int Column { get; }
            public float Left { get; }
            public float Top { get; }
            public float Width { get; }
            public float Height { get; }
        }
    }
}
