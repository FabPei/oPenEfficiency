using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnSplitByParagraphs",
        Name = "Split Text Box by Paragraphs",
        SearchAliases = new[] { "Split by Paragraphs" },
        Tooltip = "Split one text box by paragraph",
        IconData = "M3,13H2V11H3V13M3,7H2V5H3V7M3,19H2V17H3V19M21,11H5V13H21V11M21,5H5V7H21V5M21,17H5V19H21V17Z",
        Color = "#FBBF24",
        Description = "Splits one multi-paragraph text box into individually editable text boxes. Use the context menu to merge selected text boxes.",
        DetailedHelpText = "### Split Text Box by Paragraphs\n\nSelect one text box with multiple paragraphs, then split it into individually editable text boxes. Right-click to omit empty paragraphs or retain them as empty shapes. The command preserves basic font, alignment, margin, and paragraph settings where PowerPoint permits and confirms before replacing the original. **Merge Selected Text Boxes** combines multiple text boxes in top-to-bottom, left-to-right order using the basic formatting of the first shape.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class SplitByParagraphsFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, keepEmptyParagraphs: false);
        }

        public static bool Execute(PowerPointManager manager, bool keepEmptyParagraphs)
        {
            if (manager == null) return false;
            var selection = manager.GetSelectedShapes();
            if (selection == null || selection.Count != 1)
            {
                ShowInfo("Select one text box to split. To combine several text boxes, use the Merge Selected Text Boxes context-menu action.");
                return false;
            }

            var source = selection[1];
            if (source.HasTextFrame != Office.MsoTriState.msoTrue || source.TextFrame.HasText != Office.MsoTriState.msoTrue)
            {
                ShowInfo("The selected shape does not contain text.");
                return false;
            }

            var paragraphs = CollectParagraphs(source.TextFrame.TextRange, keepEmptyParagraphs);
            if (paragraphs.Count < 2)
            {
                ShowInfo(keepEmptyParagraphs
                    ? "The selected text box needs at least two paragraphs to split."
                    : "The selected text box needs at least two non-empty paragraphs to split.");
                return false;
            }

            var confirmation = System.Windows.Forms.MessageBox.Show(
                $"Split this text box into {paragraphs.Count} text boxes?\n\nEmpty paragraphs will be {(keepEmptyParagraphs ? "kept as empty shapes" : "removed")}. The original is deleted only after every new text box has been created successfully.",
                "Split Text Box by Paragraphs",
                System.Windows.Forms.MessageBoxButtons.YesNo,
                System.Windows.Forms.MessageBoxIcon.Warning);
            if (confirmation != System.Windows.Forms.DialogResult.Yes) return false;

            var created = new List<Shape>();
            try
            {
                StartUndoEntry(manager);
                float top = source.Top;
                const float spacing = 5f;
                foreach (var paragraph in paragraphs)
                {
                    var splitShape = source.Duplicate()[1];
                    splitShape.TextFrame.TextRange.Text = paragraph.Text;
                    CopyBasicTextFormatting(paragraph.Source, splitShape.TextFrame.TextRange);
                    splitShape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
                    splitShape.Top = top;
                    top += splitShape.Height + spacing;
                    created.Add(splitShape);
                }
            }
            catch (Exception ex)
            {
                DeleteShapes(created, "SplitByParagraphsFeature.Split.Rollback");
                ExceptionLogger.Log(ex, "SplitByParagraphsFeature.Split");
                ShowError("The text box could not be split. The original was left unchanged.");
                return false;
            }

            try { source.Delete(); }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SplitByParagraphsFeature.Split.DeleteOriginal");
                System.Windows.Forms.MessageBox.Show("The new text boxes were created, but the original could not be deleted and remains on the slide.", "Split Text Box by Paragraphs", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
            return true;
        }

        public static bool MergeSelectedTextBoxes(PowerPointManager manager)
        {
            if (manager == null) return false;
            var selection = manager.GetSelectedShapes();
            if (selection == null || selection.Count < 2)
            {
                ShowInfo("Select at least two text boxes to merge.");
                return false;
            }

            var shapes = new List<Shape>();
            for (int index = 1; index <= selection.Count; index++)
            {
                var shape = selection[index];
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue) shapes.Add(shape);
            }
            if (shapes.Count < 2)
            {
                ShowInfo("Select at least two text boxes containing text to merge.");
                return false;
            }

            shapes.Sort((first, second) =>
                ParagraphFlowPlanner.CompareReadingOrder(first.Top, first.Left, second.Top, second.Left));
            var confirmation = System.Windows.Forms.MessageBox.Show(
                $"Merge {shapes.Count} text boxes in reading order?\n\nThe merged text uses the basic formatting of the first shape. Originals are deleted only after the merged text box is created successfully.",
                "Merge Selected Text Boxes",
                System.Windows.Forms.MessageBoxButtons.YesNo,
                System.Windows.Forms.MessageBoxIcon.Warning);
            if (confirmation != System.Windows.Forms.DialogResult.Yes) return false;

            Shape merged = null;
            try
            {
                StartUndoEntry(manager);
                var first = shapes[0];
                var texts = new List<string>();
                foreach (var shape in shapes) texts.Add(shape.TextFrame.TextRange.Text);
                merged = first.Duplicate()[1];
                merged.TextFrame.TextRange.Text = ParagraphFlowPlanner.JoinParagraphs(texts);
                CopyBasicTextFormatting(first.TextFrame.TextRange, merged.TextFrame.TextRange);
                merged.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            }
            catch (Exception ex)
            {
                DeleteShapes(new List<Shape> { merged }, "SplitByParagraphsFeature.Merge.Rollback");
                ExceptionLogger.Log(ex, "SplitByParagraphsFeature.Merge");
                ShowError("The text boxes could not be merged. The originals were left unchanged.");
                return false;
            }

            try
            {
                foreach (var shape in shapes) shape.Delete();
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SplitByParagraphsFeature.Merge.DeleteOriginals");
                System.Windows.Forms.MessageBox.Show("The merged text box was created, but one or more originals could not be deleted. Review the slide before continuing.", "Merge Selected Text Boxes", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
            return true;
        }

        private static List<ParagraphData> CollectParagraphs(TextRange textRange, bool keepEmptyParagraphs)
        {
            var paragraphs = new List<ParagraphData>();
            int count = textRange.Paragraphs().Count;
            for (int index = 1; index <= count; index++)
            {
                var paragraph = textRange.Paragraphs(index, 1);
                var text = ParagraphFlowPlanner.WithoutParagraphEnding(paragraph.Text);
                if (ParagraphFlowPlanner.KeepParagraphForSplit(text, keepEmptyParagraphs))
                    paragraphs.Add(new ParagraphData(text, paragraph));
            }
            return paragraphs;
        }

        private static void CopyBasicTextFormatting(TextRange source, TextRange destination)
        {
            var sourceFont = source.Font;
            var destinationFont = destination.Font;
            destinationFont.Name = sourceFont.Name;
            destinationFont.Size = sourceFont.Size;
            destinationFont.Color.RGB = sourceFont.Color.RGB;
            destinationFont.Bold = sourceFont.Bold;
            destinationFont.Italic = sourceFont.Italic;
            destination.ParagraphFormat.Alignment = source.ParagraphFormat.Alignment;
        }

        private static void StartUndoEntry(PowerPointManager manager)
        {
            try { manager.GetApplication().StartNewUndoEntry(); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "SplitByParagraphsFeature.StartUndoEntry"); }
        }

        private static void DeleteShapes(List<Shape> shapes, string context)
        {
            foreach (var shape in shapes)
            {
                if (shape == null) continue;
                try { shape.Delete(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, context); }
            }
        }

        private static void ShowInfo(string message)
        {
            System.Windows.Forms.MessageBox.Show(message, "Split Text Box by Paragraphs", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        }

        private static void ShowError(string message)
        {
            System.Windows.Forms.MessageBox.Show(message, "Split Text Box by Paragraphs", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
        }

        private sealed class ParagraphData
        {
            public ParagraphData(string text, TextRange source)
            {
                Text = text;
                Source = source;
            }

            public string Text { get; }
            public TextRange Source { get; }
        }
    }
}
