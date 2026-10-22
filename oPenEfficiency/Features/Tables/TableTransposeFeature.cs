using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>Transposes a table by creating a validated replacement.</summary>
    [FeatureMetadata(
        Id = "BtnTableTranspose",
        Name = "Transpose Table",
        Tooltip = "Transpose rows and columns",
        IconData = "M3,3H21V21H3V3M7,5V19H19V5H7M9,7H17V9H9V7M9,11H17V13H9V11M9,15H14V17H9V15Z",
        Color = "#F43F5E",
        Description = "Transposes the selected table in place or creates a positioned copy while retaining its overall bounds and basic cell formatting.",
        DetailedHelpText = "### Transpose Table\n\nSwaps rows and columns in the selected table. The primary action replaces the source in place. Right-click can instead keep the original and place a transposed copy to its right or below it, clamped to the slide bounds. Relative column widths become relative row heights and vice versa. Plain text, solid fills, basic font and alignment settings, margins, and the four cell borders are copied. Rich-text runs, hyperlinks, gradients, patterns, effects, and merged cells are not supported. Merged tables are rejected without modifying the source, and any creation error leaves the original table unchanged.",
        MinSelection = 0,
        RequiresTable = true)]
    public static class TableTransposeFeature
    {
        public enum TransposeMode
        {
            ReplaceOriginal,
            CopyRight,
            CopyBelow
        }

        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, TransposeMode.ReplaceOriginal);
        }

        public static bool Execute(PowerPointManager manager, TransposeMode mode)
        {
            return ExecuteCore(manager, mode, true);
        }

        /// <summary>
        /// Runs the production mutation path without modal dialogs so an isolated Office
        /// integration harness cannot hang while reporting a failure.
        /// </summary>
        public static bool ExecuteForAutomation(PowerPointManager manager, TransposeMode mode)
        {
            return ExecuteCore(manager, mode, false);
        }

        private static bool ExecuteCore(PowerPointManager manager, TransposeMode mode, bool showMessages)
        {
            if (manager == null) return false;

            try
            {
                Shape sourceShape;
                Table sourceTable;
                if (!TableHelper.GetSelectedTable(manager, out sourceShape, out sourceTable)) return false;

                if (HasMergedCells(sourceTable))
                {
                    ShowUserMessage(showMessages, "Merged cells cannot be transposed safely. Unmerge the table cells and try again.", MessageBoxIcon.Information);
                    return false;
                }

                var app = manager.GetApplication();
                var slide = app.ActiveWindow?.View?.Slide as Slide;
                if (slide == null) return false;

                Shape resultShape = null;
                try
                {
                    try { app.StartNewUndoEntry(); }
                    catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "TableTransposeFeature.Execute.StartUndoEntry"); }

                    int sourceRows = sourceTable.Rows.Count;
                    int sourceColumns = sourceTable.Columns.Count;
                    int targetZOrder = sourceShape.ZOrderPosition;
                    string sourceName = sourceShape.Name;
                    float resultLeft = sourceShape.Left;
                    float resultTop = sourceShape.Top;
                    if (mode == TransposeMode.CopyRight)
                    {
                        float slideWidth = app.ActivePresentation.PageSetup.SlideWidth;
                        resultLeft = Math.Max(0f, Math.Min(sourceShape.Left + sourceShape.Width + 10f, slideWidth - sourceShape.Width));
                    }
                    else if (mode == TransposeMode.CopyBelow)
                    {
                        float slideHeight = app.ActivePresentation.PageSetup.SlideHeight;
                        resultTop = Math.Max(0f, Math.Min(sourceShape.Top + sourceShape.Height + 10f, slideHeight - sourceShape.Height));
                    }

                    resultShape = slide.Shapes.AddTable(sourceColumns, sourceRows, resultLeft, resultTop, sourceShape.Width, sourceShape.Height);
                    var resultTable = resultShape.Table;
                    ApplyTransposedDimensions(sourceShape, sourceTable, resultTable);

                    for (int row = 1; row <= sourceRows; row++)
                    {
                        for (int column = 1; column <= sourceColumns; column++)
                        {
                            CopyCell(sourceTable.Cell(row, column), resultTable.Cell(column, row));
                        }
                    }

                    if (mode == TransposeMode.ReplaceOriginal)
                    {
                        sourceShape.Delete();
                        try { resultShape.Name = sourceName; }
                        catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "TableTransposeFeature.Execute.RestoreName"); }
                        try { RestoreZOrder(resultShape, targetZOrder); }
                        catch (Exception ex) { ExceptionLogger.Log(ex, "TableTransposeFeature.RestoreZOrder"); }
                    }
                    else
                    {
                        try { resultShape.Name = sourceName + "_Transposed"; }
                        catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "TableTransposeFeature.Execute.CopyName"); }
                    }
                    try { resultShape.Select(); } catch (Exception ex) { ExceptionLogger.Log(ex, "TableTransposeFeature.SelectResult"); }
                    return true;
                }
                catch (Exception ex)
                {
                    TryDelete(resultShape);
                    ExceptionLogger.Log(ex, "TableTransposeFeature.CreateReplacement");
                    ShowUserMessage(showMessages, "The table could not be transposed. The original table was left unchanged.", MessageBoxIcon.Error);
                    return false;
                }
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TableTransposeFeature.Execute");
                return false;
            }
        }

        private static void ShowUserMessage(bool showMessages, string message, MessageBoxIcon icon)
        {
            if (showMessages)
                MessageBox.Show(message, "Transpose Table", MessageBoxButtons.OK, icon);
        }

        private static bool HasMergedCells(Table table)
        {
            for (int row = 1; row <= table.Rows.Count; row++)
            {
                for (int column = 1; column <= table.Columns.Count; column++)
                {
                    var shape = table.Cell(row, column).Shape;
                    if (shape.Width > table.Columns[column].Width + 0.5f ||
                        shape.Height > table.Rows[row].Height + 0.5f) return true;
                }
            }
            return false;
        }

        private static void ApplyTransposedDimensions(Shape sourceShape, Table source, Table target)
        {
            float sourceRowTotal = 0;
            for (int row = 1; row <= source.Rows.Count; row++) sourceRowTotal += source.Rows[row].Height;
            float sourceColumnTotal = 0;
            for (int column = 1; column <= source.Columns.Count; column++) sourceColumnTotal += source.Columns[column].Width;

            for (int column = 1; column <= target.Columns.Count; column++)
                target.Columns[column].Width = source.Rows[column].Height / sourceRowTotal * sourceShape.Width;
            for (int row = 1; row <= target.Rows.Count; row++)
                target.Rows[row].Height = source.Columns[row].Width / sourceColumnTotal * sourceShape.Height;
        }

        private static void CopyCell(Cell source, Cell target)
        {
            var sourceShape = source.Shape;
            var targetShape = target.Shape;
            targetShape.TextFrame.TextRange.Text = sourceShape.TextFrame.TextRange.Text;

            TryCopy(() =>
            {
                targetShape.Fill.Visible = sourceShape.Fill.Visible;
                targetShape.Fill.ForeColor.RGB = sourceShape.Fill.ForeColor.RGB;
                targetShape.Fill.Transparency = sourceShape.Fill.Transparency;
                targetShape.Fill.Solid();
            }, "TableTransposeFeature.CopyFill");

            TryCopy(() =>
            {
                var sourceFont = sourceShape.TextFrame.TextRange.Font;
                var targetFont = targetShape.TextFrame.TextRange.Font;
                targetFont.Name = sourceFont.Name;
                targetFont.Size = sourceFont.Size;
                targetFont.Color.RGB = sourceFont.Color.RGB;
                targetFont.Bold = sourceFont.Bold;
                targetFont.Italic = sourceFont.Italic;
                targetFont.Underline = sourceFont.Underline;
                targetShape.TextFrame.TextRange.ParagraphFormat.Alignment = sourceShape.TextFrame.TextRange.ParagraphFormat.Alignment;
                targetShape.TextFrame.VerticalAnchor = sourceShape.TextFrame.VerticalAnchor;
                targetShape.TextFrame.Orientation = sourceShape.TextFrame.Orientation;
                targetShape.TextFrame.MarginLeft = sourceShape.TextFrame.MarginLeft;
                targetShape.TextFrame.MarginRight = sourceShape.TextFrame.MarginRight;
                targetShape.TextFrame.MarginTop = sourceShape.TextFrame.MarginTop;
                targetShape.TextFrame.MarginBottom = sourceShape.TextFrame.MarginBottom;
            }, "TableTransposeFeature.CopyTextFormat");

            CopyBorder(source, target, PpBorderType.ppBorderTop);
            CopyBorder(source, target, PpBorderType.ppBorderBottom);
            CopyBorder(source, target, PpBorderType.ppBorderLeft);
            CopyBorder(source, target, PpBorderType.ppBorderRight);
        }

        private static void CopyBorder(Cell source, Cell target, PpBorderType borderType)
        {
            TryCopy(() =>
            {
                var sourceBorder = source.Borders[borderType];
                var targetBorder = target.Borders[borderType];
                targetBorder.Visible = sourceBorder.Visible;
                targetBorder.ForeColor.RGB = sourceBorder.ForeColor.RGB;
                targetBorder.Weight = sourceBorder.Weight;
                targetBorder.Style = sourceBorder.Style;
                targetBorder.DashStyle = sourceBorder.DashStyle;
            }, "TableTransposeFeature.CopyBorder");
        }

        private static void TryCopy(Action action, string context)
        {
            try { action(); }
            catch (Exception ex) { ExceptionLogger.Log(ex, context); }
        }

        private static void RestoreZOrder(Shape shape, int target)
        {
            int guard = 0;
            while (shape.ZOrderPosition > target && guard++ < 1000) shape.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
            while (shape.ZOrderPosition < target && guard++ < 2000) shape.ZOrder(Office.MsoZOrderCmd.msoBringForward);
        }

        private static void TryDelete(Shape shape)
        {
            if (shape == null) return;
            try { shape.Delete(); }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "TableTransposeFeature.TryDelete"); }
        }
    }
}
