using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;

namespace oPenEfficiency.Features.Text
{
    public static class TextHelper
    {
        public static bool ApplyToText(PowerPointManager manager, Action<TextRange> action)
        {
            var app = manager.GetApplication();
            if (app == null || app.ActiveWindow == null || app.ActiveWindow.Selection == null) return false;

            var sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionText)
            {
                if (sel.TextRange != null)
                {
                    action(sel.TextRange);
                    return true;
                }
            }
            else if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                bool applied = false;
                var shapes = manager.GetSelectedShapes();
                if (shapes != null)
                {
                    for (int i = 1; i <= shapes.Count; i++)
                    {
                        var shape = shapes[i];
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        {
                            action(shape.TextFrame.TextRange);
                            applied = true;
                        }
                    }
                }
                return applied;
            }
            return false;
        }

        public static bool ApplyToTextFrame(PowerPointManager manager, Action<TextFrame> action)
        {
            var app = manager.GetApplication();
            if (app == null || app.ActiveWindow == null || app.ActiveWindow.Selection == null) return false;

            var sel = app.ActiveWindow.Selection;
            if (sel.Type == PpSelectionType.ppSelectionText)
            {
                if (sel.ShapeRange.Count > 0 && sel.ShapeRange[1].HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    action(sel.ShapeRange[1].TextFrame);
                    return true;
                }
            }
            else if (sel.Type == PpSelectionType.ppSelectionShapes)
            {
                bool applied = false;
                var shapes = manager.GetSelectedShapes();
                if (shapes != null)
                {
                    for (int i = 1; i <= shapes.Count; i++)
                    {
                        var shape = shapes[i];
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            action(shape.TextFrame);
                            applied = true;
                        }
                    }
                }
                return applied;
            }
            return false;
        }

        public static Office.MsoTriState Toggle(Office.MsoTriState state)
        {
            return state == Office.MsoTriState.msoTrue ? Office.MsoTriState.msoFalse : Office.MsoTriState.msoTrue;
        }
    }

    [FeatureMetadata(Id = "BtnIncreaseFontSize", Name = "Increase Font Size", Tooltip = "Increase Font Size", IconData = "M4 19V17H20V19H4M15.4 14.1L14 10.3L12.6 14.1H15.4M14 5.3L9.5 16H11.5L12.1 14.3H15.9L16.5 16H18.5L14 5.3Z", Color = "#EAB308", Description = "Increases the font size of the selected text or shapes.", DetailedHelpText = "Increases the font size of the selected text or shapes in increments.", MinSelection = 1)]
    public static class IncreaseFontSizeFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try { manager.GetApplication().CommandBars.ExecuteMso("FontSizeIncrease"); return true; } catch { return false; }
        }
    }

    [FeatureMetadata(Id = "BtnDecreaseFontSize", Name = "Decrease Font Size", Tooltip = "Decrease Font Size", IconData = "M4 19V17H20V19H4M15.4 14.1L14 10.3L12.6 14.1H15.4M14 5.3L9.5 16H11.5L12.1 14.3H15.9L16.5 16H18.5L14 5.3Z", Color = "#EAB308", Description = "Decreases the font size of the selected text or shapes.", DetailedHelpText = "Decreases the font size of the selected text or shapes in increments.", MinSelection = 1)]
    public static class DecreaseFontSizeFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try { manager.GetApplication().CommandBars.ExecuteMso("FontSizeDecrease"); return true; } catch { return false; }
        }
    }

    [FeatureMetadata(Id = "BtnFormatBold", Name = "Bold", Tooltip = "Bold (Right-click to remove)", IconData = "M13.5 15.5H10V12.5H13.5A1.5 1.5 0 0 1 15 14A1.5 1.5 0 0 1 13.5 15.5M10 6.5H13A1.5 1.5 0 0 1 14.5 8A1.5 1.5 0 0 1 13 9.5H10V6.5M15.6 10.8A3 3 0 0 0 16 8A3 3 0 0 0 13 5H7V17H13.5A3.5 3.5 0 0 0 17 13.5A3.5 3.5 0 0 0 15.6 10.8Z", Color = "#EAB308", Description = "Applies bold formatting to selected text.", DetailedHelpText = "Applies bold formatting to selected text. Use right-click to explicitly remove bold formatting.", MinSelection = 1)]
    public static class FormatBoldFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try { manager.GetApplication().CommandBars.ExecuteMso("Bold"); return true; } catch { return false; }
        }

        public static bool Remove(PowerPointManager manager)
        {
            return TextHelper.ApplyToText(manager, tr => tr.Font.Bold = Office.MsoTriState.msoFalse);
        }
    }

    [FeatureMetadata(Id = "BtnFormatItalic", Name = "Italic", Tooltip = "Italic (Right-click to remove)", IconData = "M10 4V7H12.2L9.8 17H7V20H15V17H12.8L15.2 7H18V4H10Z", Color = "#EAB308", Description = "Applies italic formatting to selected text.", DetailedHelpText = "Applies italic formatting to selected text. Use right-click to explicitly remove italic formatting.", MinSelection = 1)]
    public static class FormatItalicFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try { manager.GetApplication().CommandBars.ExecuteMso("Italic"); return true; } catch { return false; }
        }

        public static bool Remove(PowerPointManager manager)
        {
            return TextHelper.ApplyToText(manager, tr => tr.Font.Italic = Office.MsoTriState.msoFalse);
        }
    }

    [FeatureMetadata(Id = "BtnFormatUnderline", Name = "Underline", Tooltip = "Underline (Right-click to remove)", IconData = "M12 17A6 6 0 0 0 18 11V3H15.5V11A3.5 3.5 0 0 1 12 14.5A3.5 3.5 0 0 1 8.5 11V3H6V11A6 6 0 0 0 12 17M5 19V21H19V19H5Z", Color = "#EAB308", Description = "Applies underline formatting to selected text.", DetailedHelpText = "Applies underline formatting to selected text. Use right-click to explicitly remove underline formatting.", MinSelection = 1)]
    public static class FormatUnderlineFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try { manager.GetApplication().CommandBars.ExecuteMso("Underline"); return true; } catch { return false; }
        }

        public static bool Remove(PowerPointManager manager)
        {
            return TextHelper.ApplyToText(manager, tr => tr.Font.Underline = Office.MsoTriState.msoFalse);
        }
    }

    [FeatureMetadata(Id = "BtnFormatStrikethrough", Name = "Strikethrough", Tooltip = "Strikethrough (Right-click to remove)", IconData = "M3 14H21V12H3M11 6.5C11 5.7 11.6 5 12.5 5C13.8 5 14 6.2 14 7V8H16V7C16 4.7 14.8 3 12.5 3C10.2 3 9 4.7 9 6.5C9 9 16 9.8 16 11.5V12H14V11.5C14 10.7 13.4 10 12.5 10C11.2 10 11 8.8 11 8V6.5M12.5 19C14.8 19 16 17.3 16 15.5V14H14V15.5C14 16.3 13.4 17 12.5 17C11.2 17 11 15.8 11 15V14H9V15.5C9 17.8 10.2 19 12.5 19Z", Color = "#EAB308", Description = "Applies a line through the middle of the selected text.", DetailedHelpText = "Applies a line through the middle of the selected text. Useful for indicating deleted or invalid information.", MinSelection = 1)]
    public static class FormatStrikethroughFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try { manager.GetApplication().CommandBars.ExecuteMso("Bold"); try { manager.GetApplication().CommandBars.ExecuteMso("Strikethrough"); } catch { } return true; } catch { return false; }
        }

        public static bool Remove(PowerPointManager manager)
        {
            return TextHelper.ApplyToText(manager, tr => 
            {
                try { dynamic font = tr.Font; font.StrikeThrough = Office.MsoTriState.msoFalse; } catch { }
            });
        }
    }

    [FeatureMetadata(Id = "BtnFormatSubscript", Name = "Subscript", Tooltip = "Subscript (Right-click to remove)", IconData = "M16 15V16.5H21V15L18.6 12.6C18.2 12.2 18 11.8 18 11.5C18 11 18.4 10.5 19 10.5C19.5 10.5 19.8 10.8 20 11.2L21 10.7C20.6 9.8 19.9 9.3 19 9.3C17.7 9.3 16.7 10.2 16.7 11.5C16.7 12.4 17.1 13.1 17.8 13.8L19.5 15.5V15H16M3 5V6.5L7.8 12L3 17.5V19H5.5L9 14.8L12.5 19H15V17.5L10.2 12L15 6.5V5H12.5L9 9.2L5.5 5H3Z", Color = "#EAB308", Description = "Formats selected text as subscript (smaller and below the line).", DetailedHelpText = "Formats selected text as subscript (smaller and below the line). Ideal for chemical formulas like H2O.", MinSelection = 1)]
    public static class FormatSubscriptFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try { manager.GetApplication().CommandBars.ExecuteMso("Subscript"); return true; } catch { return false; }
        }

        public static bool Remove(PowerPointManager manager)
        {
            return TextHelper.ApplyToText(manager, tr => { tr.Font.Superscript = Office.MsoTriState.msoFalse; tr.Font.Subscript = Office.MsoTriState.msoFalse; tr.Font.BaselineOffset = 0f; });
        }
    }

    [FeatureMetadata(Id = "BtnFormatSuperscript", Name = "Superscript", Tooltip = "Superscript (Right-click to remove)", IconData = "M16 11V12.5H21V11L18.6 8.6C18.2 8.2 18 7.8 18 7.5C18 7 18.4 6.5 19 6.5C19.5 6.5 19.8 6.8 20 7.2L21 6.7C20.6 5.8 19.9 5.3 19 5.3C17.7 5.3 16.7 6.2 16.7 7.5C16.7 8.4 17.1 9.1 17.8 9.8L19.5 11.5V11H16M3 5V6.5L7.8 12L3 17.5V19H5.5L9 14.8L12.5 19H15V17.5L10.2 12L15 6.5V5H12.5L9 9.2L5.5 5H3Z", Color = "#EAB308", Description = "Formats selected text as superscript (smaller and above the line).", DetailedHelpText = "Formats selected text as superscript (smaller and above the line). Essential for mathematical powers like E=mc2 or citations.", MinSelection = 1)]
    public static class FormatSuperscriptFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try { manager.GetApplication().CommandBars.ExecuteMso("Superscript"); return true; } catch { return false; }
        }

        public static bool Remove(PowerPointManager manager)
        {
            return TextHelper.ApplyToText(manager, tr => { tr.Font.Superscript = Office.MsoTriState.msoFalse; tr.Font.Subscript = Office.MsoTriState.msoFalse; tr.Font.BaselineOffset = 0f; });
        }
    }

    [FeatureMetadata(Id = "BtnChangeCase", Name = "Change Case", Tooltip = "Change Case (Right-click for options)", IconData = "M9 4V7H12L9.5 17H7.5L10 7H8V4H9M15 11V13H17.2L15.3 20H13.8L15.7 13H14V11H15Z", Color = "#EAB308", Description = "Changes the capitalization of selected text (Sentence case, lowercase, UPPERCASE, etc.).", DetailedHelpText = "Changes the capitalization of selected text. Use right-click to choose specific case types like Sentence case, lowercase, UPPERCASE, Capitalize Each Word, or tOGGLE cASE.", MinSelection = 1)]
    public static class ChangeCaseFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try { manager.GetApplication().CommandBars.ExecuteMso("ChangeCaseMenu"); return true; } catch { return false; }
        }

        public static bool ExecuteSpecific(PowerPointManager manager, PpChangeCase caseType)
        {
            return TextHelper.ApplyToText(manager, tr => tr.ChangeCase(caseType));
        }
    }

    [FeatureMetadata(Id = "BtnCharacterSpacing", Name = "Character Spacing", Tooltip = "Character Spacing (Right-click for options)", IconData = "M8 4V20H10V4H8M14 4V20H16V4H14M4 10V14H6V10H4M18 10V14H20V10H18Z", Color = "#EAB308", Description = "Adjusts the space between characters in the selected text.", DetailedHelpText = "Adjusts the space between characters in the selected text. Choose from presets like Very Tight, Tight, Normal, Loose, or Very Loose.", MinSelection = 1)]
    public static class CharacterSpacingFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try { manager.GetApplication().CommandBars.ExecuteMso("CharacterSpacingGallery"); return true; } catch { return false; }
        }

        public static bool ExecuteSpecific(PowerPointManager manager, float spacing)
        {
            return TextHelper.ApplyToText(manager, tr => 
            {
                try 
                {
                    // Access Font2 interface dynamically to set Spacing
                    dynamic font2 = tr.Font;
                    font2.Spacing = spacing;
                } 
                catch (Exception ex)
                {
                    oPenEfficiency.Utils.ExceptionLogger.Log(ex, "CharacterSpacingFeature.ExecuteSpecific");
                }
            });
        }
        
        public static bool ExecuteMoreSpacing(PowerPointManager manager)
        {
            try { manager.GetApplication().CommandBars.ExecuteMso("FontDialog"); return true; } catch { return false; }
        }
    }

    [FeatureMetadata(Id = "BtnAlignText", Name = "Align Text", Tooltip = "Align Text (Right-click for options)", IconData = "M3 3h18v2H3V3m0 8h18v2H3v-2m0 8h18v2H3v-2z", Color = "#EAB308", Description = "Adjusts the vertical and horizontal alignment of text within its container.", DetailedHelpText = "Adjusts the vertical and horizontal alignment of text within its container. Right-click to quickly set Top, Middle, or Bottom vertical alignment combined with Left, Center, or Right horizontal alignment.", MinSelection = 1)]
    public static class AlignTextFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try { manager.GetApplication().CommandBars.ExecuteMso("AlignTextGallery"); return true; } catch { return false; }
        }

        public static bool ExecuteSpecific(PowerPointManager manager, Office.MsoVerticalAnchor? anchor = null, PpParagraphAlignment? horizontalAlign = null)
        {
            return TextHelper.ApplyToTextFrame(manager, tf => 
            { 
                if (anchor.HasValue) tf.VerticalAnchor = anchor.Value; 
                if (horizontalAlign.HasValue && tf.HasText == Office.MsoTriState.msoTrue)
                {
                    tf.TextRange.ParagraphFormat.Alignment = horizontalAlign.Value;
                }
            });
        }
    }

    [FeatureMetadata(Id = "BtnTextDirection", Name = "Text Direction", Tooltip = "Text Direction (Right-click for options)", IconData = "M9 3v13.59l-3.3-3.3L4.29 14.71L10 20.41l5.71-5.7l-1.41-1.42L11 16.59V3H9m7 4v2h5V7h-5m0 4v2h5v-2h-5z", Color = "#EAB308", Description = "Rotates the text direction (Horizontal, Vertical, Stacked, etc.) within the shape.", DetailedHelpText = "Rotates the text direction within the shape. Options include Horizontal, Rotate all text 90°, Rotate all text 270°, or Stacked.", MinSelection = 1)]
    public static class TextDirectionFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try { manager.GetApplication().CommandBars.ExecuteMso("TextDirectionGallery"); return true; } catch { return false; }
        }

        public static bool ExecuteSpecific(PowerPointManager manager, Office.MsoTextOrientation orientation)
        {
            return TextHelper.ApplyToTextFrame(manager, tf => tf.Orientation = orientation);
        }
    }

    [FeatureMetadata(Id = "BtnNewSlide", Name = "New Slide", Tooltip = "New Slide (Dropdown)", IconData = "M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 14h-3v3h-2v-3H8v-2h3v-3h2v3h3v2zm-3-7V3.5L18.5 9H13z", Color = "#10B981", Description = "Inserts a new slide into the presentation using the current layout.", DetailedHelpText = "Inserts a new slide into the presentation using the current layout. Right-click to choose from available master slide layouts.", MinSelection = 0)]
    public static class NewSlideFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            try { manager.GetApplication().CommandBars.ExecuteMso("SlideNewGallery"); return true; } catch { return false; }
        }
    }
}
