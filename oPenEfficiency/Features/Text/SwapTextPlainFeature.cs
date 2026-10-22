using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;
using oPenEfficiency.Services;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Swap plain text between two shapes, ignoring formatting.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnSwapTextPlain",
        Name = "Swap Text (Plain)",
        Tooltip = "Exchange text while retaining destination formatting",
        IconData = "M19,4H15.5L14.5,3H9.5L8.5,4H5V6H19M6,19A2,2 0 0,0 8,21H16A2,2 0 0,0 18,19V7H6V19Z",
        Color = "#6B7280",
        Description = "Swaps plain text between two text shapes or selected table cells while retaining destination formatting.",
        DetailedHelpText = "### Swap Text (Plain)\n\nClick to exchange the raw text strings of exactly two selected text-bearing shapes or exactly two selected table cells. The incoming text adopts the destination formatting; source character runs are not transferred. Right-click to shrink overflowing text after the swap, cycle plain text through three or more selected targets in selection order, or choose one paragraph from each of two targets to exchange. Invalid selections show a specific explanation. Every successful action creates one PowerPoint Undo entry and attempts rollback after an error.",
        MinSelection = 1,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class SwapTextPlainFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return ExecuteSwap(manager, false);
        }

        public static bool ExecuteSwapAndShrink(PowerPointManager manager)
        {
            return ExecuteSwap(manager, true);
        }

        public static bool CycleTextThroughSelection(PowerPointManager manager)
        {
            List<TextTarget> targets;
            string selectionError;
            if (!TryGetTextTargets(manager, out targets, out selectionError))
            {
                ShowInfo(selectionError);
                return false;
            }
            if (targets.Count < 3)
            {
                ShowInfo("Select at least three text-bearing shapes or at least three table cells to cycle their text.");
                return false;
            }

            List<string> originalTexts = null;
            try
            {
                originalTexts = CaptureTexts(targets);
                StartUndoEntry(manager);
                for (int index = 0; index < targets.Count; index++)
                {
                    int sourceIndex = index == 0 ? targets.Count - 1 : index - 1;
                    targets[index].Range.Text = originalTexts[sourceIndex];
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SwapTextPlainFeature.CycleTextThroughSelection");
                if (originalTexts != null)
                    RestoreTexts(targets, originalTexts, "SwapTextPlainFeature.CycleRollback");
                ShowError("The text cycle could not be completed. The original text was restored where possible.");
                return false;
            }
        }

        public static bool SwapSelectedParagraphs(PowerPointManager manager)
        {
            List<TextTarget> targets;
            if (!TryGetExactlyTwoTargets(manager, out targets)) return false;

            int firstParagraph;
            int secondParagraph;
            if (!TryChooseParagraphs(manager, targets[0], targets[1], out firstParagraph, out secondParagraph))
                return false;

            TextRange firstRange = null;
            TextRange secondRange = null;
            string firstOriginal = null;
            string secondOriginal = null;
            try
            {
                firstRange = targets[0].Range.Paragraphs(firstParagraph, 1);
                secondRange = targets[1].Range.Paragraphs(secondParagraph, 1);
                firstOriginal = firstRange.Text;
                secondOriginal = secondRange.Text;

                StartUndoEntry(manager);
                firstRange.Text = ParagraphFlowPlanner.WithoutParagraphEnding(secondOriginal) + ParagraphFlowPlanner.ParagraphEnding(firstOriginal);
                secondRange.Text = ParagraphFlowPlanner.WithoutParagraphEnding(firstOriginal) + ParagraphFlowPlanner.ParagraphEnding(secondOriginal);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SwapTextPlainFeature.SwapSelectedParagraphs");
                try
                {
                    if (firstRange != null && firstOriginal != null) firstRange.Text = firstOriginal;
                    if (secondRange != null && secondOriginal != null) secondRange.Text = secondOriginal;
                }
                catch (Exception rollbackEx)
                {
                    ExceptionLogger.Log(rollbackEx, "SwapTextPlainFeature.ParagraphRollback");
                }
                ShowError("The selected paragraphs could not be exchanged. The original text was restored where possible.");
                return false;
            }
        }

        private static bool ExecuteSwap(PowerPointManager manager, bool shrinkOverflow)
        {
            List<TextTarget> targets;
            if (!TryGetExactlyTwoTargets(manager, out targets)) return false;

            List<string> originalTexts = null;
            List<Office.MsoAutoSize?> originalAutoSizes = null;
            try
            {
                originalTexts = CaptureTexts(targets);
                if (shrinkOverflow) originalAutoSizes = CaptureAutoSizes(targets);
                StartUndoEntry(manager);
                targets[0].Range.Text = originalTexts[1];
                targets[1].Range.Text = originalTexts[0];

                if (shrinkOverflow)
                {
                    ShrinkIfOverflowing(targets[0].Shape);
                    ShrinkIfOverflowing(targets[1].Shape);
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SwapTextPlainFeature.ExecuteSwap");
                if (originalTexts != null)
                    RestoreTexts(targets, originalTexts, "SwapTextPlainFeature.SwapRollback");
                if (originalAutoSizes != null)
                    RestoreAutoSizes(targets, originalAutoSizes);
                ShowError("The text could not be exchanged. The original text was restored where possible.");
                return false;
            }
        }

        private static bool TryGetExactlyTwoTargets(PowerPointManager manager, out List<TextTarget> targets)
        {
            string selectionError;
            if (!TryGetTextTargets(manager, out targets, out selectionError))
            {
                ShowInfo(selectionError);
                return false;
            }
            if (targets.Count != 2)
            {
                ShowInfo("Select exactly two text-bearing shapes or exactly two table cells.");
                return false;
            }
            return true;
        }

        private static bool TryGetTextTargets(
            PowerPointManager manager,
            out List<TextTarget> targets,
            out string error)
        {
            targets = new List<TextTarget>();
            error = null;
            if (manager == null)
            {
                error = "PowerPoint is not available.";
                return false;
            }

            try
            {
                var shapes = manager.GetSelectedShapes();
                if (shapes == null || shapes.Count == 0)
                {
                    error = "Select text-bearing shapes or individual table cells first.";
                    return false;
                }

                if (shapes.Count == 1 && shapes[1].HasTable == Office.MsoTriState.msoTrue)
                {
                    var table = shapes[1].Table;
                    for (int row = 1; row <= table.Rows.Count; row++)
                    {
                        for (int column = 1; column <= table.Columns.Count; column++)
                        {
                            var cell = table.Cell(row, column);
                            if (!cell.Selected) continue;
                            var cellShape = cell.Shape;
                            if (cellShape.HasTextFrame == Office.MsoTriState.msoTrue)
                            {
                                targets.Add(new TextTarget(
                                    cellShape,
                                    cellShape.TextFrame.TextRange,
                                    string.Format("Cell {0}, {1}", row, column)));
                            }
                        }
                    }

                    if (targets.Count == 0)
                    {
                        error = "Select individual table cells. Selecting the complete table is not supported.";
                        return false;
                    }
                    return true;
                }

                for (int index = 1; index <= shapes.Count; index++)
                {
                    var shape = shapes[index];
                    if (shape.HasTable == Office.MsoTriState.msoTrue)
                    {
                        error = "Select individual cells inside one table, or select only regular text-bearing shapes.";
                        targets.Clear();
                        return false;
                    }
                    if (shape.HasTextFrame != Office.MsoTriState.msoTrue)
                    {
                        error = string.Format("The selected shape \"{0}\" does not contain editable text.", shape.Name);
                        targets.Clear();
                        return false;
                    }
                    targets.Add(new TextTarget(shape, shape.TextFrame.TextRange, shape.Name));
                }
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SwapTextPlainFeature.TryGetTextTargets");
                targets.Clear();
                error = "The current PowerPoint selection could not be read.";
                return false;
            }
        }

        private static List<string> CaptureTexts(List<TextTarget> targets)
        {
            var texts = new List<string>(targets.Count);
            foreach (var target in targets) texts.Add(target.Range.Text);
            return texts;
        }

        private static List<Office.MsoAutoSize?> CaptureAutoSizes(List<TextTarget> targets)
        {
            var states = new List<Office.MsoAutoSize?>(targets.Count);
            foreach (var target in targets)
            {
                try { states.Add(target.Shape.TextFrame2.AutoSize); }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "SwapTextPlainFeature.CaptureAutoSize");
                    states.Add(null);
                }
            }
            return states;
        }

        private static void RestoreTexts(List<TextTarget> targets, List<string> texts, string context)
        {
            for (int index = 0; index < targets.Count && index < texts.Count; index++)
            {
                try { targets[index].Range.Text = texts[index]; }
                catch (Exception ex) { ExceptionLogger.Log(ex, context); }
            }
        }

        private static void RestoreAutoSizes(List<TextTarget> targets, List<Office.MsoAutoSize?> states)
        {
            for (int index = 0; index < targets.Count && index < states.Count; index++)
            {
                if (!states[index].HasValue) continue;
                try { targets[index].Shape.TextFrame2.AutoSize = states[index].Value; }
                catch (Exception ex) { ExceptionLogger.Log(ex, "SwapTextPlainFeature.AutoSizeRollback"); }
            }
        }

        private static void ShrinkIfOverflowing(Shape shape)
        {
            var frame = shape.TextFrame2;
            float availableWidth = Math.Max(0f, shape.Width - frame.MarginLeft - frame.MarginRight);
            float availableHeight = Math.Max(0f, shape.Height - frame.MarginTop - frame.MarginBottom);
            bool verticalOverflow = frame.TextRange.BoundHeight > availableHeight + 0.5f;
            bool horizontalOverflow =
                frame.WordWrap == Office.MsoTriState.msoFalse &&
                frame.TextRange.BoundWidth > availableWidth + 0.5f;
            if (verticalOverflow || horizontalOverflow)
                frame.AutoSize = Office.MsoAutoSize.msoAutoSizeTextToFitShape;
        }

        private static bool TryChooseParagraphs(
            PowerPointManager manager,
            TextTarget first,
            TextTarget second,
            out int firstParagraph,
            out int secondParagraph)
        {
            firstParagraph = 0;
            secondParagraph = 0;
            var firstChoices = GetParagraphChoices(first.Range);
            var secondChoices = GetParagraphChoices(second.Range);
            if (firstChoices.Count == 0 || secondChoices.Count == 0)
            {
                ShowInfo("Both selected targets need at least one paragraph.");
                return false;
            }

            var dialog = new ParagraphSwapDialog(first.Label, firstChoices, second.Label, secondChoices);
            try
            {
                var app = manager.GetApplication();
                if (app != null)
                    new WindowInteropHelper(dialog).Owner = (IntPtr)app.HWND;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "SwapTextPlainFeature.SetParagraphDialogOwner");
            }

            if (dialog.ShowDialog() != true) return false;
            firstParagraph = dialog.FirstParagraph;
            secondParagraph = dialog.SecondParagraph;
            return true;
        }

        private static List<ParagraphChoice> GetParagraphChoices(TextRange range)
        {
            var choices = new List<ParagraphChoice>();
            int count = range.Paragraphs().Count;
            for (int index = 1; index <= count; index++)
            {
                choices.Add(new ParagraphChoice(
                    index,
                    ParagraphFlowPlanner.DescribeParagraphChoice(index, range.Paragraphs(index, 1).Text)));
            }
            return choices;
        }

        private static void StartUndoEntry(PowerPointManager manager)
        {
            try { manager.GetApplication()?.StartNewUndoEntry(); }
            catch (Exception ex) { ExceptionLogger.Log(ex, "SwapTextPlainFeature.StartNewUndoEntry"); }
        }

        private static void ShowInfo(string message)
        {
            System.Windows.Forms.MessageBox.Show(
                message,
                "Swap Text (Plain)",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Information);
        }

        private static void ShowError(string message)
        {
            System.Windows.Forms.MessageBox.Show(
                message,
                "Swap Text (Plain)",
                System.Windows.Forms.MessageBoxButtons.OK,
                System.Windows.Forms.MessageBoxIcon.Error);
        }

        private sealed class TextTarget
        {
            public TextTarget(Shape shape, TextRange range, string label)
            {
                Shape = shape;
                Range = range;
                Label = label;
            }

            public Shape Shape { get; private set; }
            public TextRange Range { get; private set; }
            public string Label { get; private set; }
        }

        private sealed class ParagraphChoice
        {
            public ParagraphChoice(int index, string display)
            {
                Index = index;
                Display = display;
            }

            public int Index { get; private set; }
            public string Display { get; private set; }
        }

        private sealed class ParagraphSwapDialog : Window
        {
            private readonly ComboBox _first;
            private readonly ComboBox _second;

            public ParagraphSwapDialog(
                string firstLabel,
                List<ParagraphChoice> firstChoices,
                string secondLabel,
                List<ParagraphChoice> secondChoices)
            {
                Title = "Swap Selected Paragraphs";
                Width = 620;
                SizeToContent = SizeToContent.Height;
                ResizeMode = ResizeMode.NoResize;
                WindowStartupLocation = WindowStartupLocation.CenterOwner;
                ShowInTaskbar = false;

                var root = new Grid { Margin = new Thickness(18) };
                for (int index = 0; index < 6; index++)
                    root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                AddText(root, 0, "Choose one paragraph from each selected target. Only the paragraph text is exchanged; each destination keeps its formatting.", FontWeights.Normal, null);
                AddText(root, 1, firstLabel, FontWeights.SemiBold, new Thickness(0, 14, 0, 5));
                _first = AddComboBox(root, 2, firstChoices);
                AddText(root, 3, secondLabel, FontWeights.SemiBold, new Thickness(0, 14, 0, 5));
                _second = AddComboBox(root, 4, secondChoices);

                var buttons = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    Margin = new Thickness(0, 18, 0, 0)
                };
                var cancel = new Button
                {
                    Content = "Cancel",
                    IsCancel = true,
                    MinWidth = 90,
                    Margin = new Thickness(0, 0, 8, 0)
                };
                var swap = new Button
                {
                    Content = "Swap Paragraphs",
                    IsDefault = true,
                    MinWidth = 130
                };
                swap.Click += delegate { DialogResult = true; };
                buttons.Children.Add(cancel);
                buttons.Children.Add(swap);
                Grid.SetRow(buttons, 5);
                root.Children.Add(buttons);
                Content = root;
            }

            public int FirstParagraph
            {
                get { return ((ParagraphChoice)_first.SelectedItem).Index; }
            }

            public int SecondParagraph
            {
                get { return ((ParagraphChoice)_second.SelectedItem).Index; }
            }

            private static void AddText(
                Grid root,
                int row,
                string text,
                FontWeight weight,
                Thickness? margin)
            {
                var block = new TextBlock
                {
                    Text = text,
                    TextWrapping = TextWrapping.Wrap,
                    FontWeight = weight,
                    Margin = margin ?? new Thickness(0)
                };
                Grid.SetRow(block, row);
                root.Children.Add(block);
            }

            private static ComboBox AddComboBox(Grid root, int row, List<ParagraphChoice> choices)
            {
                var combo = new ComboBox
                {
                    ItemsSource = choices,
                    DisplayMemberPath = "Display",
                    SelectedIndex = 0,
                    MinWidth = 560
                };
                Grid.SetRow(combo, row);
                root.Children.Add(combo);
                return combo;
            }
        }
    }
}
