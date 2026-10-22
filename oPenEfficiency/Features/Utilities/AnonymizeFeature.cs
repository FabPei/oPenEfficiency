using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    [FeatureMetadata(
        Id = "BtnAnonymize",
        Name = "Anonymize Selected Slide Text",
        Tooltip = "Replace text on selected slides with placeholders",
        IconData = "M12,2A10,10 0 0,1 22,12A10,10 0 0,1 12,22A10,10 0 0,1 2,12A10,10 0 0,1 12,2M12,4A8,8 0 0,0 4,12A8,8 0 0,0 12,20A8,8 0 0,0 20,12A8,8 0 0,0 12,4M12,6A6,6 0 0,1 18,12A6,6 0 0,1 12,18A6,6 0 0,1 6,12A6,6 0 0,1 12,6Z",
        Color = "#6B7280",
        Description = "Replaces visible text on selected slides with placeholder words while retaining text-frame structure and character formatting where PowerPoint permits. Dynamic slide numbers are preserved.",
        DetailedHelpText = "### Anonymize Selected Slide Text\n\nReplaces visible text on the current slide, selected slides, or the entire presentation. Right-click selects the scope or previews the affected word count without changing content. The command processes ordinary text boxes, grouped shapes, and table cells while retaining whitespace, punctuation, letter case, and character formatting where PowerPoint permits. It preserves dynamic slide-number placeholders and asks for confirmation before replacement.\n\nThis command does not remove speaker notes, comments, document properties, alternative text, chart data, or embedded-file metadata. Remove those items separately when required.",
        MinSelection = 0)]
    public static class AnonymizeFeature
    {
        public enum AnonymizeScope { SelectedSlides, CurrentSlide, EntirePresentation }

        private static readonly Random Random = new Random();
        private static readonly object RandomLock = new object();

        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, AnonymizeScope.SelectedSlides, false);
        }

        public static bool Execute(PowerPointManager manager, AnonymizeScope scope, bool previewOnly)
        {
            if (manager == null) return false;
            try
            {
                var slides = GetTargetSlides(manager, scope);
                if (slides.Count == 0)
                {
                    System.Windows.MessageBox.Show(
                        "Select one or more slides, or select a shape on the slide whose visible text you want to replace.",
                        "Anonymize Selected Slide Text",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Information);
                    return false;
                }

                var replacements = new List<TextReplacement>();
                foreach (Slide slide in slides)
                {
                    foreach (Shape shape in slide.Shapes)
                    {
                        CollectReplacements(shape, replacements);
                    }
                }

                if (replacements.Count == 0)
                {
                    System.Windows.MessageBox.Show(
                        "No visible text was found on the selected slide scope.",
                        "Anonymize Selected Slide Text",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Information);
                    return false;
                }

                if (previewOnly)
                {
                    System.Windows.MessageBox.Show(
                        $"{replacements.Count} word(s) on {slides.Count} slide(s) would be anonymized.\n\nNo changes were made.",
                        "Anonymization Preview",
                        System.Windows.MessageBoxButton.OK,
                        System.Windows.MessageBoxImage.Information);
                    return true;
                }

                var confirmation = System.Windows.MessageBox.Show(
                    $"Replace {replacements.Count} word(s) on {slides.Count} selected slide(s)?\n\n" +
                    "Text formatting and paragraph structure are retained where PowerPoint permits. " +
                    "Speaker notes, comments, document properties, alternative text, chart data, and embedded-file metadata are not changed.",
                    "Anonymize Selected Slide Text",
                    System.Windows.MessageBoxButton.YesNo,
                    System.Windows.MessageBoxImage.Warning);
                if (confirmation != System.Windows.MessageBoxResult.Yes) return false;

                try { manager.GetApplication().StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "AnonymizeFeature.StartNewUndoEntry"); }

                int applied = 0;
                try
                {
                    foreach (var replacement in replacements)
                    {
                        replacement.Range.Text = replacement.Replacement;
                        applied++;
                    }
                }
                catch
                {
                    for (int i = applied - 1; i >= 0; i--)
                    {
                        try { replacements[i].Range.Text = replacements[i].OriginalText; }
                        catch (Exception rollbackEx) { ExceptionLogger.Log(rollbackEx, "AnonymizeFeature.Execute.Rollback"); }
                    }
                    throw;
                }

                System.Windows.MessageBox.Show(
                    $"Replaced {replacements.Count} word(s) on {slides.Count} slide(s).",
                    "Anonymize Selected Slide Text",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Information);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "AnonymizeFeature.Execute");
                return false;
            }
        }

        private static List<Slide> GetTargetSlides(PowerPointManager manager, AnonymizeScope scope)
        {
            var result = new List<Slide>();
            var app = manager?.GetApplication();
            if (app?.ActivePresentation == null) return result;

            if (scope == AnonymizeScope.EntirePresentation)
            {
                foreach (Slide slide in app.ActivePresentation.Slides) result.Add(slide);
                return result;
            }

            if (scope == AnonymizeScope.CurrentSlide)
            {
                var slide = app.ActiveWindow?.View?.Slide as Slide;
                if (slide != null) result.Add(slide);
                return result;
            }

            return manager.GetSelectedSlidesForOperation();
        }

        private static void CollectReplacements(Shape shape, List<TextReplacement> replacements)
        {
            if (IsSlideNumberPlaceholder(shape)) return;

            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                foreach (Shape child in shape.GroupItems) CollectReplacements(child, replacements);
                return;
            }

            if (shape.HasTable == Office.MsoTriState.msoTrue)
            {
                var table = shape.Table;
                for (int row = 1; row <= table.Rows.Count; row++)
                {
                    for (int column = 1; column <= table.Columns.Count; column++)
                    {
                        CollectTextRangeReplacements(table.Cell(row, column).Shape.TextFrame.TextRange, replacements);
                    }
                }
                return;
            }

            if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
            {
                CollectTextRangeReplacements(shape.TextFrame.TextRange, replacements);
            }
        }

        private static bool IsSlideNumberPlaceholder(Shape shape)
        {
            return shape.Type == Office.MsoShapeType.msoPlaceholder &&
                   shape.PlaceholderFormat.Type == PpPlaceholderType.ppPlaceholderSlideNumber;
        }

        private static void CollectTextRangeReplacements(TextRange textRange, List<TextReplacement> replacements)
        {
            int count = textRange.Words().Count;
            for (int index = 1; index <= count; index++)
            {
                var word = textRange.Words(index, 1);
                var original = word.Text;
                if (!string.IsNullOrWhiteSpace(original))
                {
                    replacements.Add(new TextReplacement(word, original, GeneratePlaceholder(original)));
                }
            }
        }

        private static string GeneratePlaceholder(string original)
        {
            lock (RandomLock)
            {
                return oPenEfficiency.Core.Services.AnonymizePlaceholderPolicy.GeneratePlaceholder(original, Random);
            }
        }

        private sealed class TextReplacement
        {
            public TextReplacement(TextRange range, string originalText, string replacement)
            {
                Range = range;
                OriginalText = originalText;
                Replacement = replacement;
            }

            public TextRange Range { get; }
            public string OriginalText { get; }
            public string Replacement { get; }
        }
    }
}
