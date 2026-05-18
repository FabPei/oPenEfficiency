using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features
{
    /// <summary>
    /// Feature: Split a textbox into multiple shapes by paragraph, or merge multiple shapes into one.
    /// </summary>
    [FeatureMetadata(
        Id = "BtnSplitByParagraphs",
        Name = "Split by Paragraphs",
        Tooltip = "Split by paragraphs",
        IconData = "M3,13H2V11H3V13M3,7H2V5H3V7M3,19H2V17H3V19M21,11H5V13H21V11M21,5H5V7H21V5M21,17H5V19H21V17Z",
        Color = "#FBBF24",
        Description = "Splits a textbox into multiple shapes by paragraph, or merges multiple shapes into one.",
        DetailedHelpText = "### Split by Paragraphs\nSplits a multi-paragraph text box into individual separate text boxes, one per paragraph, while preserving font, size, and color of each line.",
        MinSelection = 1,
        RequiredType = PpSelectionType.ppSelectionShapes)]
    public static class SplitByParagraphsFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            if (manager == null) return false;
            var app = manager.GetApplication();
            var activeWindow = app.ActiveWindow;
            if (activeWindow == null) return false;

            var sel = activeWindow.Selection;
            if (sel == null) return false;

            // Handle Text Selection (Partial Split)
            if (sel.Type == PpSelectionType.ppSelectionText)
            {
                try
                {
                    var textRange = sel.TextRange;
                    var shape = sel.ShapeRange[1];
                    int count = textRange.Paragraphs().Count;
                    if (count == 0) return false;

                    float top = textRange.BoundTop;
                    float spacing = 5f;

                    // Collect paragraph texts first
                    var texts = new List<string>();
                    for (int i = 1; i <= count; i++)
                    {
                        string t = textRange.Paragraphs(i, 1).Text.TrimEnd('\r', '\n');
                        if (!string.IsNullOrEmpty(t)) texts.Add(t);
                    }

                    if (texts.Count == 0) return false;

                    // Delete the selected text from original shape
                    textRange.Delete();

                    // Create new shapes for each selected paragraph
                    foreach (var t in texts)
                    {
                        var newShape = shape.Duplicate()[1];
                        newShape.TextFrame.TextRange.Text = t;
                        newShape.Top = top;
                        top += newShape.Height + spacing;
                    }

                    return true;
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "SplitByParagraphsFeature.Execute.textSelection");
                    return false;
                }
            }

            // Normal Shape Selection Logic
            var selection = manager.GetSelectedShapes();
            if (selection == null || selection.Count == 0) return false;

            if (selection.Count == 1)
            {
                var shape = selection[1];
                if (shape.HasTextFrame != Office.MsoTriState.msoTrue) return false;

                try
                {
                    var textRange = shape.TextFrame.TextRange;
                    int count = textRange.Paragraphs().Count;
                    if (count <= 1)
                    {
                        System.Windows.Forms.MessageBox.Show(
                            "The selected shape has only one paragraph or no text.\\n\\n" +
                            "Split by Paragraphs requires multiple paragraphs to split.",
                            "Split by Paragraphs",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Information);
                        return true;
                    }

                    float top = shape.Top;
                    float spacing = 5f; // Small gap between split objects

                    for (int i = 1; i <= count; i++)
                    {
                        var para = textRange.Paragraphs(i, 1);
                        var newShape = shape.Duplicate()[1];
                        newShape.Top = top;
                        newShape.TextFrame.TextRange.Text = para.Text.TrimEnd('\r', '\n');
                        top += newShape.Height + spacing;
                    }

                    shape.Delete();
                    return true;
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "SplitByParagraphsFeature.Execute.singleShape");
                    return false;
                }
            }
            else
            {
                // Merge Mode
                try
                {
                    var shapes = new List<Shape>();
                    for (int i = 1; i <= selection.Count; i++)
                    {
                        if (selection[i].HasTextFrame == Office.MsoTriState.msoTrue)
                            shapes.Add(selection[i]);
                    }

                    if (shapes.Count <= 1)
                    {
                        System.Windows.Forms.MessageBox.Show(
                            "Not enough shapes with text to merge.\\n\\n" +
                            "Please select multiple shapes to merge their text.",
                            "Split by Paragraphs",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Information);
                        return false;
                    }

                    // Sort by Vertical then Horizontal
                    shapes.Sort((a, b) =>
                    {
                        int res = a.Top.CompareTo(b.Top);
                        if (Math.Abs(a.Top - b.Top) < 5) res = a.Left.CompareTo(b.Left);
                        return res;
                    });

                    var first = shapes[0];
                    string combinedText = "";
                    foreach (var s in shapes)
                    {
                        string t = s.TextFrame.TextRange.Text.TrimEnd('\r', '\n');
                        if (!string.IsNullOrEmpty(t))
                            combinedText += t + "\r";
                    }

                    var newShape = first.Duplicate()[1];
                    newShape.TextFrame.TextRange.Text = combinedText.TrimEnd('\r');
                    
                    foreach (var s in shapes) { try { s.Delete(); } catch (Exception ex) { ExceptionLogger.Log(ex, "SplitByParagraphsFeature.Merge.Delete"); } }

                    return true;
                }
                catch (Exception ex)
                {
                    ExceptionLogger.Log(ex, "SplitByParagraphsFeature.Execute.mergeMode");
                    return false;
                }
            }
        }
    }
}
