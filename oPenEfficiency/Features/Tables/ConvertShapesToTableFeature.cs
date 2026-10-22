using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>Converts a uniform grid of shapes into a PowerPoint table.</summary>
    [FeatureMetadata(
        Id = "BtnConvertToTable",
        Name = "Convert Shapes to Table",
        Tooltip = "Convert a shape grid to a table",
        IconData = "M3,3H21V21H3V3M5,5V19H19V5H5M7,7H17V9H7V7M7,11H17V13H7V11M7,15H13V17H7V15Z",
        Color = "#F43F5E",
        Description = "Converts selected, uniformly aligned shapes into a PowerPoint table while retaining text and available basic cell formatting.",
        DetailedHelpText = "### Convert Shapes to Table\n\nSelect two or more non-overlapping shapes arranged in a uniform grid. Rectangles and rounded rectangles work especially well, but other shapes are accepted when you only need their text and available basic formatting in table cells. The command validates row and column alignment before it creates a table. The original shapes remain unchanged until conversion succeeds, then you can delete them, keep them visible, or discard the new table.\n\nThe conversion retains plain text plus available basic fill, line, font, alignment, and margin settings. Shapes without text create empty cells. Rich text, hyperlinks, gradients, patterns, effects, images, custom geometry, rotations, and gaps between shapes are not retained. Review the table before deleting the original shapes.",
        MinSelection = 2,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class ConvertShapesToTableFeature
    {
        public enum OriginalShapeDisposition
        {
            AskAfterConversion,
            KeepOriginals,
            ReplaceOriginals
        }

        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, OriginalShapeDisposition.AskAfterConversion);
        }

        /// <param name="interactive">
        /// False selects the message-box-free programmatic path (acceptance seam): the conversion
        /// proceeds without the confirmation prompt, AskAfterConversion skips its follow-up
        /// question (keeping the original shapes unless the disposition says ReplaceOriginals),
        /// and no message box is ever raised. Default is true, so the dialog-driven user path is
        /// unchanged.
        /// </param>
        public static bool Execute(
            PowerPointManager manager,
            OriginalShapeDisposition disposition,
            bool interactive = true)
        {
            if (manager == null) return false;

            try
            {
                var app = manager.GetApplication();
                if (app.ActiveWindow == null) return false;

                var shapes = GetSelectedShapes(manager, interactive);
                if (shapes == null) return false;

                var rows = BuildUniformGrid(shapes);
                if (rows == null)
                {
                    if (interactive)
                    {
                        MessageBox.Show("Select shapes arranged in a uniform grid with the same number of columns in every row.", "Convert Shapes to Table", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    return false;
                }

                if (interactive)
                {
                    var confirmation = MessageBox.Show(
                        "Convert the selected shape grid into a table?\n\n" +
                        "The original shapes remain unchanged until conversion succeeds. Text and available basic formatting are copied; shapes without text create empty cells. Images, custom geometry, rotations, rich text, hyperlinks, gradients, effects, and gaps between shapes are not retained.",
                        "Convert Shapes to Table",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning);
                    if (confirmation != DialogResult.Yes) return false;
                }

                var slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                if (slide == null) return false;

                PowerPoint.Shape tableShape = null;
                try
                {
                    try { app.StartNewUndoEntry(); } catch { /* Conversion remains valid when undo grouping is unavailable. */ }
                    tableShape = CreateTable(slide, rows);
                }
                catch (Exception ex)
                {
                    DeleteTable(tableShape);
                    ExceptionLogger.Log(ex, "ConvertShapesToTableFeature.Convert");
                    if (interactive)
                    {
                        MessageBox.Show("The table could not be created. The original shapes were left unchanged.", "Convert Shapes to Table", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    return false;
                }

                DialogResult result;
                if (disposition == OriginalShapeDisposition.AskAfterConversion)
                {
                    // Non-interactive keeps the original shapes (same default as the table-side
                    // conversion path); ReplaceOriginals/KeepOriginals never reach this dialog.
                    result = interactive
                        ? MessageBox.Show(
                            "Conversion succeeded.\n\nYes: delete the original shapes\nNo: keep the original shapes visible\nCancel: discard the new table and keep the original shapes",
                            "Convert Shapes to Table",
                            MessageBoxButtons.YesNoCancel,
                            MessageBoxIcon.Question)
                        : DialogResult.No;
                }
                else
                {
                    result = disposition == OriginalShapeDisposition.ReplaceOriginals
                        ? DialogResult.Yes
                        : DialogResult.No;
                }
                if (result == DialogResult.Cancel)
                {
                    DeleteTable(tableShape);
                    return false;
                }

                if (result == DialogResult.Yes)
                {
                    try
                    {
                        foreach (var shape in shapes) shape.Delete();
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "ConvertShapesToTableFeature.DeleteOriginals");
                        if (interactive)
                        {
                            MessageBox.Show("The table was created, but one or more original shapes could not be deleted. Review the slide before continuing.", "Convert Shapes to Table", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }

                try { tableShape.Select(); } catch (Exception ex) { ExceptionLogger.Log(ex, "ConvertShapesToTableFeature.SelectResult"); }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ConvertShapesToTableFeature.Execute");
                if (interactive)
                {
                    MessageBox.Show("An unexpected error occurred. No original shapes were intentionally removed.", "Convert Shapes to Table", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return false;
            }
        }

        private static List<PowerPoint.Shape> GetSelectedShapes(PowerPointManager manager, bool interactive)
        {
            var selection = manager.GetSelectedShapes();
            if (selection == null || selection.Count < 2)
            {
                if (interactive)
                {
                    MessageBox.Show("Select at least two shapes to convert to a table.", "Convert Shapes to Table", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                return null;
            }

            var shapes = new List<PowerPoint.Shape>();
            for (int index = 1; index <= selection.Count; index++)
            {
                shapes.Add(selection[index]);
            }
            return shapes;
        }

        private static List<List<PowerPoint.Shape>> BuildUniformGrid(List<PowerPoint.Shape> shapes)
        {
            if (shapes.Any(shape => shape.Width <= 0 || shape.Height <= 0)) return null;

            float minimumHeight = shapes.Min(shape => shape.Height);
            float minimumWidth = shapes.Min(shape => shape.Width);
            float rowTolerance = Math.Max(2f, Math.Min(12f, minimumHeight * 0.15f));
            float columnTolerance = Math.Max(2f, Math.Min(12f, minimumWidth * 0.15f));

            var rows = new List<List<PowerPoint.Shape>>();
            foreach (var shape in shapes.OrderBy(shape => shape.Top))
            {
                var row = rows.FirstOrDefault(existing => Math.Abs(existing[0].Top - shape.Top) <= rowTolerance);
                if (row == null)
                {
                    row = new List<PowerPoint.Shape>();
                    rows.Add(row);
                }
                row.Add(shape);
            }

            foreach (var row in rows) row.Sort((first, second) => first.Left.CompareTo(second.Left));
            rows.Sort((first, second) => first[0].Top.CompareTo(second[0].Top));

            int columnCount = rows[0].Count;
            if (columnCount == 0 || rows.Any(row => row.Count != columnCount)) return null;

            for (int column = 0; column < columnCount; column++)
            {
                float expectedLeft = rows[0][column].Left;
                float expectedWidth = rows[0][column].Width;
                if (rows.Any(row => Math.Abs(row[column].Left - expectedLeft) > columnTolerance || Math.Abs(row[column].Width - expectedWidth) > columnTolerance)) return null;
            }

            foreach (var row in rows)
            {
                float expectedHeight = row[0].Height;
                if (row.Any(shape => Math.Abs(shape.Height - expectedHeight) > rowTolerance)) return null;
                for (int column = 1; column < row.Count; column++)
                    if (row[column].Left < row[column - 1].Left + row[column - 1].Width - columnTolerance) return null;
            }
            for (int row = 1; row < rows.Count; row++)
                if (rows[row][0].Top < rows[row - 1][0].Top + rows[row - 1].Max(shape => shape.Height) - rowTolerance) return null;
            return rows;
        }

        private static PowerPoint.Shape CreateTable(PowerPoint.Slide slide, List<List<PowerPoint.Shape>> rows)
        {
            int rowCount = rows.Count;
            int columnCount = rows[0].Count;
            var columnWidths = new float[columnCount];
            var rowHeights = new float[rowCount];
            for (int column = 0; column < columnCount; column++) columnWidths[column] = rows.Max(row => row[column].Width);
            for (int row = 0; row < rowCount; row++) rowHeights[row] = rows[row].Max(shape => shape.Height);

            float left = rows[0][0].Left;
            float top = rows[0][0].Top;
            float width = columnWidths.Sum();
            float height = rowHeights.Sum();
            var tableShape = slide.Shapes.AddTable(rowCount, columnCount, left, top, width, height);
            var table = tableShape.Table;

            for (int column = 0; column < columnCount; column++) table.Columns[column + 1].Width = columnWidths[column];
            for (int row = 0; row < rowCount; row++) table.Rows[row + 1].Height = rowHeights[row];

            for (int row = 0; row < rowCount; row++)
            {
                for (int column = 0; column < columnCount; column++) CopyBasicAppearance(rows[row][column], table.Cell(row + 1, column + 1).Shape);
            }
            return tableShape;
        }

        private static void CopyBasicAppearance(PowerPoint.Shape source, PowerPoint.Shape destination)
        {
            try
            {
                destination.Fill.Visible = source.Fill.Visible;
                destination.Fill.ForeColor.RGB = source.Fill.ForeColor.RGB;
                destination.Fill.Transparency = source.Fill.Transparency;
                destination.Fill.Solid();
            }
            catch (Exception ex) { ExceptionLogger.Log(ex, "ConvertShapesToTableFeature.CopyFill"); }
            try
            {
                destination.Line.Visible = source.Line.Visible;
                destination.Line.ForeColor.RGB = source.Line.ForeColor.RGB;
                destination.Line.Weight = source.Line.Weight;
                destination.Line.DashStyle = source.Line.DashStyle;
            }
            catch (Exception ex)
            {
                // Table cells do not expose the Line object at all (Shape.get_Line throws
                // NotImplementedException on a cell), so this capability gap is expected on every
                // conversion and is neither logged nor treated as a failure; real line-copy
                // failures still reach the error log.
                if (!(ex is NotImplementedException))
                {
                    ExceptionLogger.Log(ex, "ConvertShapesToTableFeature.CopyLine");
                }
            }

            if (source.HasTextFrame != Office.MsoTriState.msoTrue || source.TextFrame.HasText != Office.MsoTriState.msoTrue) return;
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
            destination.TextFrame.MarginLeft = source.TextFrame.MarginLeft;
            destination.TextFrame.MarginRight = source.TextFrame.MarginRight;
            destination.TextFrame.MarginTop = source.TextFrame.MarginTop;
            destination.TextFrame.MarginBottom = source.TextFrame.MarginBottom;
        }

        private static void DeleteTable(PowerPoint.Shape tableShape)
        {
            if (tableShape == null) return;
            try { tableShape.Delete(); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "ConvertShapesToTableFeature.DeleteTable"); }
        }
    }
}
