using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Swap text between two shapes while preserving their formatting.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnSwapTextFormatted",
        Name = "Swap Text (Formatted)",
        Tooltip = "Swap text (formatted)",
        IconData = "M19,4H15.5L14.5,3H9.5L8.5,4H5V6H19M6,19A2,2 0 0,0 8,21H16A2,2 0 0,0 18,19V7H6V19Z",
        Color = "#6B7280",
        Description = "Swaps text and its character formatting between exactly two selected shapes.",
        DetailedHelpText = "### Swap Text (Formatted)\n\nExchanges the complete text ranges of exactly two selected text-bearing shapes. Font runs, paragraph formatting, and other text formatting travel with the source text; shape geometry and shape formatting remain in place. The operation uses PowerPoint's clipboard, cleans up its temporary slide objects, attempts to restore both text ranges after an error, and creates one Undo entry.",
        MinSelection = 2,
        MaxSelection = 2,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class SwapTextFormattedFeature
    {
        public enum FitMode
        {
            None,
            ShrinkOverflow,
            FitShapesToText
        }

        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, FitMode.None);
        }

        public static bool Execute(PowerPointManager manager, FitMode fitMode)
        {
            if (manager == null) return false;
            var shapes = manager.GetSelectedShapes();
            if (shapes == null || shapes.Count != 2) return false;

            Shape temp1 = null;
            Shape temp2 = null;
            bool originalsCaptured = false;
            try
            {
                var s1 = shapes[1];
                var s2 = shapes[2];
                if (s1.HasTextFrame != Office.MsoTriState.msoTrue || s2.HasTextFrame != Office.MsoTriState.msoTrue) return false;

                var slide = s1.Parent as Slide;
                if (slide == null) return false;
                try { manager.GetApplication()?.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "SwapTextFormattedFeature.StartNewUndoEntry"); }

                temp1 = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, -1000, -1000, 10, 10);
                temp2 = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, -1000, -1000, 10, 10);

                s1.TextFrame.TextRange.Copy();
                temp1.TextFrame.TextRange.Paste();
                s2.TextFrame.TextRange.Copy();
                temp2.TextFrame.TextRange.Paste();
                originalsCaptured = true;

                temp2.TextFrame.TextRange.Copy();
                s1.TextFrame.TextRange.Paste();

                temp1.TextFrame.TextRange.Copy();
                s2.TextFrame.TextRange.Paste();

                ApplyFitMode(s1, fitMode);
                ApplyFitMode(s2, fitMode);

                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SwapTextFormattedFeature.Execute");
                if (originalsCaptured)
                {
                    try
                    {
                        temp1.TextFrame.TextRange.Copy();
                        shapes[1].TextFrame.TextRange.Paste();
                        temp2.TextFrame.TextRange.Copy();
                        shapes[2].TextFrame.TextRange.Paste();
                    }
                    catch (Exception rollbackEx)
                    {
                        ExceptionLogger.Log(rollbackEx, "SwapTextFormattedFeature.Rollback");
                    }
                }
                return false;
            }
            finally
            {
                try { temp1?.Delete(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "SwapTextFormattedFeature.DeleteTemp1"); }
                try { temp2?.Delete(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "SwapTextFormattedFeature.DeleteTemp2"); }
            }
        }

        private static void ApplyFitMode(Shape shape, FitMode fitMode)
        {
            if (shape == null || shape.HasTextFrame != Office.MsoTriState.msoTrue) return;
            if (fitMode == FitMode.FitShapesToText)
            {
                shape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            }
            else if (fitMode == FitMode.ShrinkOverflow)
            {
                try
                {
                    shape.TextFrame2.AutoSize = Office.MsoAutoSize.msoAutoSizeTextToFitShape;
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "SwapTextFormattedFeature.ApplyFitMode");
                }
            }
        }
    }
}
