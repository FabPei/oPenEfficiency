using System;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Text
{
    internal static class ParagraphSpacingContextActions
    {
        public static bool OpenDialog(PowerPointManager manager)
        {
            try
            {
                var app = manager?.GetApplication();
                if (app?.CommandBars == null) return false;
                app.CommandBars.ExecuteMso("ParagraphDialog");
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ParagraphSpacingContextActions.OpenDialog");
                return false;
            }
        }

        public static bool ApplyLineSpacing(PowerPointManager manager, float lines)
        {
            return Apply(manager, paragraph =>
            {
                paragraph.LineRuleWithin = MsoTriState.msoTrue;
                paragraph.SpaceWithin = lines;
            });
        }

        public static bool RemoveParagraphSpacing(PowerPointManager manager)
        {
            return Apply(manager, paragraph =>
            {
                paragraph.LineRuleBefore = MsoTriState.msoFalse;
                paragraph.LineRuleAfter = MsoTriState.msoFalse;
                paragraph.SpaceBefore = 0f;
                paragraph.SpaceAfter = 0f;
            });
        }

        private static bool Apply(PowerPointManager manager, Action<ParagraphFormat> action)
        {
            try
            {
                var app = manager?.GetApplication();
                var selection = app?.ActiveWindow?.Selection;
                if (selection == null) return false;
                try { app.StartNewUndoEntry(); } catch { /* Paragraph changes remain valid when undo grouping is unavailable. */ }

                bool changed = false;
                if (selection.Type == PpSelectionType.ppSelectionText)
                {
                    action(selection.TextRange.ParagraphFormat);
                    return true;
                }

                if (selection.Type != PpSelectionType.ppSelectionShapes) return false;
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in selection.ShapeRange)
                {
                    if (shape.HasTable == MsoTriState.msoTrue)
                    {
                        for (int row = 1; row <= shape.Table.Rows.Count; row++)
                        for (int column = 1; column <= shape.Table.Columns.Count; column++)
                        {
                            var cellShape = shape.Table.Cell(row, column).Shape;
                            if (cellShape.TextFrame.HasText == MsoTriState.msoTrue)
                            {
                                action(cellShape.TextFrame.TextRange.ParagraphFormat);
                                changed = true;
                            }
                        }
                    }
                    else if (shape.HasTextFrame == MsoTriState.msoTrue &&
                             shape.TextFrame.HasText == MsoTriState.msoTrue)
                    {
                        action(shape.TextFrame.TextRange.ParagraphFormat);
                        changed = true;
                    }
                }
                return changed;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "ParagraphSpacingContextActions.Apply");
                return false;
            }
        }
    }
}
