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
        Name = "Split / Merge Text",
        Tooltip = "Split shape at cursor, by paragraphs, or merge multiple shapes",
        IconData = "M3,13H2V11H3V13M3,7H2V5H3V7M3,19H2V17H3V19M21,11H5V13H21V11M21,5H5V7H21V5M21,17H5V19H21V17Z",
        Color = "#FBBF24",
        Description = "Splits a shape at the cursor, by paragraphs, or merges multiple shapes into one.",
        DetailedHelpText = "### Split by Paragraphs\n**Split at cursor:** Click inside a text box to place the cursor, then press Split — the shape is divided at that exact position into two shapes (text before cursor / text after cursor).\n\n**Split selected paragraphs:** Select one or more paragraphs in a text box — each selected paragraph is extracted into its own new shape.\n\n**Split all paragraphs:** Select the whole shape (not in edit mode) with one shape — every paragraph becomes a separate shape.\n\n**Merge:** Select multiple shapes — their text is merged into one shape sorted top-to-bottom, left-to-right.",
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

            // Handle Text Edit Mode (cursor blinking or text selected)
            if (sel.Type == PpSelectionType.ppSelectionText)
            {
                var textRange = sel.TextRange;
                var shape = sel.ShapeRange[1];

                // Cursor blinking with no selection — split at exact cursor position
                if (textRange.Length == 0)
                {
                    try
                    {
                        int cursorStart = textRange.Start;
                        int fullLen = shape.TextFrame.TextRange.Length;

                        // Nothing meaningful to split if cursor is at the very start or end
                        if (cursorStart <= 1 || cursorStart > fullLen) return false;

                        float originalLeft = shape.Left;
                        float originalTop = shape.Top;
                        const float spacing = 5f;

                        // Duplicate for the "after" part (Duplicate() offsets by ~10pt, so fix position)
                        var shapeAfter = shape.Duplicate()[1];
                        shapeAfter.Left = originalLeft;
                        shapeAfter.Top = originalTop;

                        // Remove text from cursorStart onward in original (keep "before")
                        int afterLen = fullLen - cursorStart + 1;
                        if (afterLen > 0)
                            shape.TextFrame.TextRange.Characters(cursorStart, afterLen).Delete();

                        // Remove text before cursorStart in duplicate (keep "after")
                        int beforeLen = cursorStart - 1;
                        if (beforeLen > 0)
                            shapeAfter.TextFrame.TextRange.Characters(1, beforeLen).Delete();

                        // Stack duplicate directly below original (which may have shrunk)
                        shapeAfter.Top = shape.Top + shape.Height + spacing;

                        return true;
                    }
                    catch (Exception ex)
                    {
                        ExceptionLogger.Log(ex, "SplitByParagraphsFeature.Execute.splitAtCursor", showErrorUI: true);
                        return false;
                    }
                }

                // Text selected — extract each selected paragraph into its own shape
                try
                {
                    int count = textRange.Paragraphs().Count;
                    if (count == 0) return false;

                    float top = textRange.BoundTop;
                    const float spacing = 5f;

                    // Collect paragraph texts before mutating the range
                    var texts = new List<string>();
                    for (int i = 1; i <= count; i++)
                    {
                        string t = textRange.Paragraphs(i, 1).Text.TrimEnd('\r', '\n');
                        if (!string.IsNullOrEmpty(t)) texts.Add(t);
                    }

                    if (texts.Count == 0) return false;

                    textRange.Delete();

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

                    // Use each paragraph's visual BoundTop so shapes land exactly
                    // where the text was, not stacked with artificial gaps.
                    float marginTop;
                    try { marginTop = shape.TextFrame.MarginTop; }
                    catch { marginTop = 0f; }

                    float fallbackTop = shape.Top;
                    const float fallbackSpacing = 5f;

                    for (int i = 1; i <= count; i++)
                    {
                        var para = textRange.Paragraphs(i, 1);
                        string paraText = para.Text.TrimEnd('\r', '\n');

                        float paraTop;
                        try { paraTop = para.BoundTop - marginTop; }
                        catch { paraTop = fallbackTop; }

                        var newShape = shape.Duplicate()[1];
                        newShape.Top = paraTop;
                        newShape.TextFrame.TextRange.Text = paraText;

                        fallbackTop = paraTop + newShape.Height + fallbackSpacing;
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
