using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Clear text content from all selected shapes while preserving formatting.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnDeleteText",
        Name = "Delete Text",
        Tooltip = "Clear text while keeping the selected shapes",
        IconData = "M19,4H15.5L14.5,3H9.5L8.5,4H5V6H19M6,19A2,2 0 0,0 8,21H16A2,2 0 0,0 18,19V7H6V19Z",
        Color = "#EF4444",
        Description = "Clears text from selected shapes or selected table cells while preserving the shapes themselves.",
        DetailedHelpText = "### Delete Text\n\nClears text from selected text-bearing shapes, selected table cells, and supported children of selected groups while keeping the shapes. Right-click to delete all text, only highlighted characters, or the paragraphs intersecting the active text selection. **Delete Highlighted Text** and **Delete Selected Paragraphs** are available only for an actual text selection. You can also open **Replace Selected Text...**. Each deletion creates one PowerPoint Undo entry.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class DeleteTextFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            var shapes = manager.GetSelectedShapes();
            if (shapes == null || shapes.Count == 0) return false;

            try
            {
                try { manager.GetApplication()?.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "DeleteTextFeature.StartNewUndoEntry"); }

                bool modified = false;
                bool failed = false;
                for (int i = 1; i <= shapes.Count; i++)
                {
                    modified |= ClearShapeText(shapes[i], ref failed);
                }
                return modified && !failed;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DeleteTextFeature.Execute");
                return false;
            }
        }

        public static bool HasHighlightedText(PowerPointManager manager)
        {
            try
            {
                var selection = manager?.GetApplication()?.ActiveWindow?.Selection;
                return selection != null &&
                       selection.Type == PpSelectionType.ppSelectionText &&
                       selection.TextRange != null &&
                       selection.TextRange.Length > 0;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DeleteTextFeature.HasHighlightedText");
                return false;
            }
        }

        public static bool DeleteHighlightedText(PowerPointManager manager)
        {
            return DeleteActiveTextRange(manager, selectedParagraphs: false);
        }

        public static bool DeleteSelectedParagraphs(PowerPointManager manager)
        {
            return DeleteActiveTextRange(manager, selectedParagraphs: true);
        }

        private static bool DeleteActiveTextRange(PowerPointManager manager, bool selectedParagraphs)
        {
            if (!HasHighlightedText(manager)) return false;
            try
            {
                var application = manager.GetApplication();
                try { application.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "DeleteTextFeature.DeleteActiveTextRange.StartNewUndoEntry"); }

                var textRange = application.ActiveWindow.Selection.TextRange;
                if (selectedParagraphs) textRange = textRange.Paragraphs();
                textRange.Text = string.Empty;
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "DeleteTextFeature.DeleteActiveTextRange");
                return false;
            }
        }

        private static bool ClearShapeText(Shape shape, ref bool failed)
        {
            bool modified = false;
            try
            {
                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    for (int i = 1; i <= shape.GroupItems.Count; i++)
                        modified |= ClearShapeText(shape.GroupItems[i], ref failed);
                    return modified;
                }

                if (shape.HasTable == Office.MsoTriState.msoTrue)
                {
                    bool hasSelectedCells = false;
                    for (int row = 1; row <= shape.Table.Rows.Count && !hasSelectedCells; row++)
                    {
                        for (int column = 1; column <= shape.Table.Columns.Count; column++)
                        {
                            if (shape.Table.Cell(row, column).Selected)
                            {
                                hasSelectedCells = true;
                                break;
                            }
                        }
                    }

                    for (int row = 1; row <= shape.Table.Rows.Count; row++)
                    {
                        for (int column = 1; column <= shape.Table.Columns.Count; column++)
                        {
                            var cell = shape.Table.Cell(row, column);
                            if (hasSelectedCells && !cell.Selected) continue;

                            var cellShape = cell.Shape;
                            if (cellShape.HasTextFrame == Office.MsoTriState.msoTrue &&
                                cellShape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                            {
                                cellShape.TextFrame.TextRange.Text = "";
                                modified = true;
                            }
                        }
                    }
                    return modified;
                }

                if (shape.HasTextFrame == Office.MsoTriState.msoTrue &&
                    shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    shape.TextFrame.TextRange.Text = "";
                    modified = true;
                }
            }
            catch (Exception ex)
            {
                failed = true;
                ExceptionLogger.Log(ex, "DeleteTextFeature.ClearShapeText");
            }
            return modified;
        }
    }
}
