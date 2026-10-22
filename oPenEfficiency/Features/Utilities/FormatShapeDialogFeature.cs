using System;
using Microsoft.Office.Interop.PowerPoint;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Utilities
{
    [FeatureMetadata(
        Id = "BtnFormatShapeDialog",
        Name = "Open Format Shape Pane",
        SearchAliases = new[] { "Format Shape Dialog" },
        Tooltip = "Open Format Shape",
        IconData = "M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V5h14v14z M7 7h10v2H7V7zm0 4h10v2H7v-2zm0 4h7v2H7v-2z",
        Color = "#8B5CF6",
        Description = "Opens the native PowerPoint 'Format Shape' pane to adjust fill, line, and effects.",
        DetailedHelpText = "### Open Format Shape Pane\n\nOpens PowerPoint's native Format Shape interface for the current shape selection. Right-click can remove individual effects or all effects. Reset to Default Shape Formatting uses a newly inserted AutoShape of the same type as the theme-derived reference while preserving the selected shape's geometry, position, text, and text formatting. Reset is unavailable for objects without a meaningful same-type AutoShape default.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class FormatShapeDialogFeature
    {
        public enum ContextAction
        {
            ResetToDefault,
            RemoveShadow,
            RemoveReflection,
            RemoveGlowAndSoftEdges,
            RemoveThreeD,
            RemoveAllEffects
        }

        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            try
            {
                var app = manager.GetApplication();
                var selection = app?.ActiveWindow?.Selection;
                if (selection == null || selection.Type != PpSelectionType.ppSelectionShapes)
                    return false;

                if (app.CommandBars != null && app.CommandBars.GetEnabledMso("ObjectFormatDialog"))
                {
                    app.CommandBars.ExecuteMso("ObjectFormatDialog");
                    return true;
                }

                System.Windows.Forms.MessageBox.Show(
                    "PowerPoint's Format Shape command is not available for the current selection.",
                    "Open Format Shape Pane",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information);
                return false;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FormatShapeDialogFeature.Execute");
                System.Windows.Forms.MessageBox.Show(
                    "PowerPoint could not open the Format Shape interface.",
                    "Open Format Shape Pane",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
                return false;
            }
        }

        public static bool Execute(PowerPointManager manager, ContextAction action)
        {
            var shapes = manager?.GetSelectedShapes();
            if (shapes == null || shapes.Count == 0) return false;
            try
            {
                try { manager.GetApplication()?.StartNewUndoEntry(); } catch { /* Formatting remains valid when undo grouping is unavailable. */ }
                bool changed = false;
                foreach (Shape shape in shapes)
                {
                    switch (action)
                    {
                        case ContextAction.ResetToDefault:
                            changed |= ResetToDefault(shape);
                            break;
                        case ContextAction.RemoveShadow:
                            shape.Shadow.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                            changed = true;
                            break;
                        case ContextAction.RemoveReflection:
                            shape.Reflection.Type = Microsoft.Office.Core.MsoReflectionType.msoReflectionTypeNone;
                            changed = true;
                            break;
                        case ContextAction.RemoveGlowAndSoftEdges:
                            shape.Glow.Radius = 0f;
                            shape.SoftEdge.Radius = 0f;
                            changed = true;
                            break;
                        case ContextAction.RemoveThreeD:
                            shape.ThreeD.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                            changed = true;
                            break;
                        case ContextAction.RemoveAllEffects:
                            shape.Reflection.Type = Microsoft.Office.Core.MsoReflectionType.msoReflectionTypeNone;
                            shape.Glow.Radius = 0f;
                            shape.SoftEdge.Radius = 0f;
                            shape.ThreeD.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                            // The shadow has to be cleared LAST: clearing glow/soft edges re-resolves
                            // the shape's effect style and can re-apply the theme shadow, which would
                            // survive a shadow-first ordering (QA-S9 batch 13 leg 11, host-verified).
                            shape.Shadow.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
                            changed = true;
                            break;
                    }
                }
                return changed;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "FormatShapeDialogFeature.ExecuteContextAction");
                return false;
            }
        }

        private static bool ResetToDefault(Shape shape)
        {
            if (shape == null || shape.Type != Microsoft.Office.Core.MsoShapeType.msoAutoShape) return false;
            var slide = shape.Parent as Slide;
            if (slide == null) return false;

            Shape reference = null;
            Shape textBackup = null;
            try
            {
                if (shape.HasTextFrame == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    textBackup = slide.Shapes.AddTextbox(
                        Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal,
                        -2000f, -2000f, Math.Max(10f, shape.Width), Math.Max(10f, shape.Height));
                    shape.TextFrame.TextRange.Copy();
                    textBackup.TextFrame.TextRange.Paste();
                }

                reference = slide.Shapes.AddShape(shape.AutoShapeType, -2000f, -2000f, 20f, 20f);
                reference.PickUp();
                shape.Apply();

                if (textBackup != null)
                {
                    textBackup.TextFrame.TextRange.Copy();
                    shape.TextFrame.TextRange.Paste();
                }
                return true;
            }
            finally
            {
                try { reference?.Delete(); } catch { /* Temporary reference shape may already have been removed. */ }
                try { textBackup?.Delete(); } catch { /* Temporary text backup may already have been removed. */ }
            }
        }
    }
}
