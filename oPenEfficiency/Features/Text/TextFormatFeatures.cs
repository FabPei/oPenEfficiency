using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Text
{
    public static class TextHelper
    {
        public static bool ExecuteNativeCommand(PowerPointManager manager, string commandId, string context)
        {
            if (manager == null) return false;
            try
            {
                var commandBars = manager.GetApplication()?.CommandBars;
                if (commandBars == null || !commandBars.GetEnabledMso(commandId))
                    return false;

                commandBars.ExecuteMso(commandId);
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, context);
                return false;
            }
        }

        public static bool ApplyToText(PowerPointManager manager, Action<TextRange> action)
        {
            if (manager == null || action == null) return false;
            try
            {
                var app = manager.GetApplication();
                if (app == null || app.ActiveWindow == null || app.ActiveWindow.Selection == null)
                    return false;

                var ranges = new System.Collections.Generic.List<TextRange>();
                var sel = app.ActiveWindow.Selection;
                if (sel.Type == PpSelectionType.ppSelectionText)
                {
                    if (sel.TextRange != null) ranges.Add(sel.TextRange);
                }
                else if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                    var shapes = manager.GetSelectedShapes();
                    if (shapes != null)
                    {
                        for (int i = 1; i <= shapes.Count; i++)
                        {
                            var shape = shapes[i];
                            if (shape.HasTextFrame == Office.MsoTriState.msoTrue &&
                                shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                                ranges.Add(shape.TextFrame.TextRange);
                        }
                    }
                }

                if (ranges.Count == 0) return false;
                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "TextHelper.ApplyToText.StartNewUndoEntry"); }

                bool applied = false;
                bool failed = false;
                foreach (var range in ranges)
                {
                    try
                    {
                        action(range);
                        applied = true;
                    }
                    catch (Exception ex)
                    {
                        failed = true;
                        ExceptionLogger.Log(ex, "TextHelper.ApplyToText.Range");
                    }
                }
                return applied && !failed;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TextHelper.ApplyToText");
                return false;
            }
        }

        public static bool ApplyToTextFrame(PowerPointManager manager, Action<TextFrame> action, string context)
        {
            if (manager == null || action == null) return false;

            try
            {
                var app = manager.GetApplication();
                if (app == null || app.ActiveWindow == null || app.ActiveWindow.Selection == null)
                    return false;

                var frames = new System.Collections.Generic.List<TextFrame>();
                var sel = app.ActiveWindow.Selection;
                if (sel.Type == PpSelectionType.ppSelectionText)
                {
                    if (sel.ShapeRange.Count > 0 &&
                        sel.ShapeRange[1].HasTextFrame == Office.MsoTriState.msoTrue)
                        frames.Add(sel.ShapeRange[1].TextFrame);
                }
                else if (sel.Type == PpSelectionType.ppSelectionShapes)
                {
                    var shapes = manager.GetSelectedShapes();
                    if (shapes != null)
                    {
                        for (int i = 1; i <= shapes.Count; i++)
                        {
                            var shape = shapes[i];
                            if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                                frames.Add(shape.TextFrame);
                        }
                    }
                }

                if (frames.Count == 0) return false;
                try { app.StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, context + ".StartNewUndoEntry"); }

                bool applied = false;
                bool failed = false;
                foreach (var frame in frames)
                {
                    try
                    {
                        action(frame);
                        applied = true;
                    }
                    catch (Exception ex)
                    {
                        failed = true;
                        ExceptionLogger.Log(ex, context + ".TextFrame");
                    }
                }
                return applied && !failed;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, context);
                return false;
            }
        }

        public static Office.MsoTriState Toggle(Office.MsoTriState state)
        {
            return state == Office.MsoTriState.msoTrue ? Office.MsoTriState.msoFalse : Office.MsoTriState.msoTrue;
        }
    }

    [FeatureMetadata(Id = "BtnIncreaseFontSize", Name = "Increase Font Size", Tooltip = "Increase font size using PowerPoint's next step", IconData = "M4 19V17H20V19H4M15.4 14.1L14 10.3L12.6 14.1H15.4M14 5.3L9.5 16H11.5L12.1 14.3H15.9L16.5 16H18.5L14 5.3Z", Color = "#EAB308", Description = "Moves selected text to PowerPoint's next larger font-size step.", DetailedHelpText = "### Increase Font Size\n\nUses PowerPoint's native Increase Font Size command. Right-click to move one, two, or five steps on PowerPoint's built-in size ladder. It affects selected characters or text in selected text-bearing shapes; mixed font sizes remain relative.", MinSelection = 1)]
    public static class IncreaseFontSizeFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return TextHelper.ExecuteNativeCommand(manager, "FontSizeIncrease", "IncreaseFontSizeFeature.Execute");
        }

        public static bool ExecuteSteps(PowerPointManager manager, int steps)
        {
            bool changed = false;
            for (int step = 0; step < steps; step++) changed |= Execute(manager);
            return changed;
        }
    }

    [FeatureMetadata(Id = "BtnDecreaseFontSize", Name = "Decrease Font Size", Tooltip = "Decrease font size using PowerPoint's previous step", IconData = "M4 19V17H20V19H4M15.4 14.1L14 10.3L12.6 14.1H15.4M14 5.3L9.5 16H11.5L12.1 14.3H15.9L16.5 16H18.5L14 5.3Z", Color = "#EAB308", Description = "Moves selected text to PowerPoint's next smaller font-size step.", DetailedHelpText = "### Decrease Font Size\n\nUses PowerPoint's native Decrease Font Size command. Right-click to move one, two, or five steps on PowerPoint's built-in size ladder. It affects selected characters or text in selected text-bearing shapes; mixed font sizes remain relative.", MinSelection = 1)]
    public static class DecreaseFontSizeFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return TextHelper.ExecuteNativeCommand(manager, "FontSizeDecrease", "DecreaseFontSizeFeature.Execute");
        }

        public static bool ExecuteSteps(PowerPointManager manager, int steps)
        {
            bool changed = false;
            for (int step = 0; step < steps; step++) changed |= Execute(manager);
            return changed;
        }
    }

    [FeatureMetadata(Id = "BtnFormatBold", Name = "Bold", Tooltip = "Toggle bold formatting for selected text", IconData = "M13.5 15.5H10V12.5H13.5A1.5 1.5 0 0 1 15 14A1.5 1.5 0 0 1 13.5 15.5M10 6.5H13A1.5 1.5 0 0 1 14.5 8A1.5 1.5 0 0 1 13 9.5H10V6.5M15.6 10.8A3 3 0 0 0 16 8A3 3 0 0 0 13 5H7V17H13.5A3.5 3.5 0 0 0 17 13.5A3.5 3.5 0 0 0 15.6 10.8Z", Color = "#EAB308", Description = "Toggles PowerPoint bold formatting for the current text selection.", DetailedHelpText = "### Bold\n\nThe primary action runs PowerPoint's native Bold toggle. Right-click to explicitly apply, remove, or toggle Bold for selected characters or all text in selected text-bearing shapes. Shapes without text are skipped and each explicit action supports Undo.", MinSelection = 1)]
    public static class FormatBoldFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return TextHelper.ExecuteNativeCommand(manager, "Bold", "FormatBoldFeature.Execute");
        }

        public static bool Remove(PowerPointManager manager)
        {
            return TextHelper.ApplyToText(manager, tr => tr.Font.Bold = Office.MsoTriState.msoFalse);
        }

        public static bool Apply(PowerPointManager manager)
        {
            return TextHelper.ApplyToText(manager, tr => tr.Font.Bold = Office.MsoTriState.msoTrue);
        }
    }

    [FeatureMetadata(Id = "BtnFormatItalic", Name = "Italic", Tooltip = "Toggle italic formatting for selected text", IconData = "M10 4V7H12.2L9.8 17H7V20H15V17H12.8L15.2 7H18V4H10Z", Color = "#EAB308", Description = "Toggles PowerPoint italic formatting for the current text selection.", DetailedHelpText = "### Italic\n\nThe primary action runs PowerPoint's native Italic toggle. Right-click to explicitly apply, remove, or toggle Italic for selected characters or all text in selected text-bearing shapes. Shapes without text are skipped and each explicit action supports Undo.", MinSelection = 1)]
    public static class FormatItalicFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return TextHelper.ExecuteNativeCommand(manager, "Italic", "FormatItalicFeature.Execute");
        }

        public static bool Remove(PowerPointManager manager)
        {
            return TextHelper.ApplyToText(manager, tr => tr.Font.Italic = Office.MsoTriState.msoFalse);
        }

        public static bool Apply(PowerPointManager manager)
        {
            return TextHelper.ApplyToText(manager, tr => tr.Font.Italic = Office.MsoTriState.msoTrue);
        }
    }

    [FeatureMetadata(Id = "BtnFormatUnderline", Name = "Underline", Tooltip = "Toggle underline formatting for selected text", IconData = "M12 17A6 6 0 0 0 18 11V3H15.5V11A3.5 3.5 0 0 1 12 14.5A3.5 3.5 0 0 1 8.5 11V3H6V11A6 6 0 0 0 12 17M5 19V21H19V19H5Z", Color = "#EAB308", Description = "Toggles PowerPoint underline formatting for the current text selection.", DetailedHelpText = "### Underline\n\nThe primary action runs PowerPoint's native Underline toggle. Right-click to explicitly apply, remove, or toggle Underline for selected characters or all text in selected text-bearing shapes. Shapes without text are skipped and each explicit action supports Undo.", MinSelection = 1)]
    public static class FormatUnderlineFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return TextHelper.ExecuteNativeCommand(manager, "Underline", "FormatUnderlineFeature.Execute");
        }

        public static bool Remove(PowerPointManager manager)
        {
            return TextHelper.ApplyToText(manager, tr => tr.Font.Underline = Office.MsoTriState.msoFalse);
        }

        public static bool Apply(PowerPointManager manager)
        {
            return TextHelper.ApplyToText(manager, tr => tr.Font.Underline = Office.MsoTriState.msoTrue);
        }
    }

    [FeatureMetadata(Id = "BtnFormatStrikethrough", Name = "Strikethrough", Tooltip = "Toggle strikethrough for selected text", IconData = "M3 14H21V12H3M11 6.5C11 5.7 11.6 5 12.5 5C13.8 5 14 6.2 14 7V8H16V7C16 4.7 14.8 3 12.5 3C10.2 3 9 4.7 9 6.5C9 9 16 9.8 16 11.5V12H14V11.5C14 10.7 13.4 10 12.5 10C11.2 10 11 8.8 11 8V6.5M12.5 19C14.8 19 16 17.3 16 15.5V14H14V15.5C14 16.3 13.4 17 12.5 17C11.2 17 11 15.8 11 15V14H9V15.5C9 17.8 10.2 19 12.5 19Z", Color = "#EAB308", Description = "Toggles PowerPoint strikethrough without changing bold formatting.", DetailedHelpText = "### Strikethrough\n\nRuns PowerPoint's native Strikethrough toggle on the current text selection. PowerPoint's Font object model exposes no strikethrough property, so oPenEfficiency cannot explicitly apply or remove strikethrough: each click flips the selection's state and PowerPoint Undo reverts it. Right-click repeats the toggle.", MinSelection = 1)]
    public static class FormatStrikethroughFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return TextHelper.ExecuteNativeCommand(manager, "Strikethrough", "FormatStrikethroughFeature.Execute");
        }

    }

    [FeatureMetadata(Id = "BtnFormatSubscript", Name = "Subscript", Tooltip = "Toggle subscript for selected text", IconData = "M16 15V16.5H21V15L18.6 12.6C18.2 12.2 18 11.8 18 11.5C18 11 18.4 10.5 19 10.5C19.5 10.5 19.8 10.8 20 11.2L21 10.7C20.6 9.8 19.9 9.3 19 9.3C17.7 9.3 16.7 10.2 16.7 11.5C16.7 12.4 17.1 13.1 17.8 13.8L19.5 15.5V15H16M3 5V6.5L7.8 12L3 17.5V19H5.5L9 14.8L12.5 19H15V17.5L10.2 12L15 6.5V5H12.5L9 9.2L5.5 5H3Z", Color = "#EAB308", Description = "Toggles PowerPoint subscript for the current text selection.", DetailedHelpText = "### Subscript\n\nThe primary action runs PowerPoint's native Subscript toggle. Right-click to explicitly apply, remove, or toggle Subscript for selected characters or all text in selected text-bearing shapes. PowerPoint keeps it mutually exclusive with Superscript; each explicit action supports Undo.", MinSelection = 1)]
    public static class FormatSubscriptFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return TextHelper.ExecuteNativeCommand(manager, "Subscript", "FormatSubscriptFeature.Execute");
        }

        public static bool Remove(PowerPointManager manager)
        {
            return TextHelper.ApplyToText(manager, tr => tr.Font.Subscript = Office.MsoTriState.msoFalse);
        }

        public static bool Apply(PowerPointManager manager)
        {
            return TextHelper.ApplyToText(manager, tr => tr.Font.Subscript = Office.MsoTriState.msoTrue);
        }
    }

    [FeatureMetadata(Id = "BtnFormatSuperscript", Name = "Superscript", Tooltip = "Toggle superscript for selected text", IconData = "M16 11V12.5H21V11L18.6 8.6C18.2 8.2 18 7.8 18 7.5C18 7 18.4 6.5 19 6.5C19.5 6.5 19.8 6.8 20 7.2L21 6.7C20.6 5.8 19.9 5.3 19 5.3C17.7 5.3 16.7 6.2 16.7 7.5C16.7 8.4 17.1 9.1 17.8 9.8L19.5 11.5V11H16M3 5V6.5L7.8 12L3 17.5V19H5.5L9 14.8L12.5 19H15V17.5L10.2 12L15 6.5V5H12.5L9 9.2L5.5 5H3Z", Color = "#EAB308", Description = "Toggles PowerPoint superscript for the current text selection.", DetailedHelpText = "### Superscript\n\nThe primary action runs PowerPoint's native Superscript toggle. Right-click to explicitly apply, remove, or toggle Superscript for selected characters or all text in selected text-bearing shapes. PowerPoint keeps it mutually exclusive with Subscript; each explicit action supports Undo.", MinSelection = 1)]
    public static class FormatSuperscriptFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return TextHelper.ExecuteNativeCommand(manager, "Superscript", "FormatSuperscriptFeature.Execute");
        }

        public static bool Remove(PowerPointManager manager)
        {
            return TextHelper.ApplyToText(manager, tr => tr.Font.Superscript = Office.MsoTriState.msoFalse);
        }

        public static bool Apply(PowerPointManager manager)
        {
            return TextHelper.ApplyToText(manager, tr => tr.Font.Superscript = Office.MsoTriState.msoTrue);
        }
    }

    [FeatureMetadata(Id = "BtnChangeCase", Name = "Change Case", Tooltip = "Open PowerPoint's case choices", IconData = "M9 4V7H12L9.5 17H7.5L10 7H8V4H9M15 11V13H17.2L15.3 20H13.8L15.7 13H14V11H15Z", Color = "#EAB308", Description = "Opens PowerPoint's Change Case menu or applies a selected case directly.", DetailedHelpText = "### Change Case\n\nThe primary action opens PowerPoint's native Change Case menu. Right-click to directly apply Sentence case, lowercase, UPPERCASE, Capitalize Each Word, or tOGGLE cASE to selected characters or all text in selected text-bearing shapes. PowerPoint controls language-specific casing; non-letter characters and formatting remain unchanged, and direct choices create one Undo entry.", MinSelection = 1)]
    public static class ChangeCaseFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return TextHelper.ExecuteNativeCommand(manager, "ChangeCaseMenu", "ChangeCaseFeature.Execute");
        }

        public static bool ExecuteSpecific(PowerPointManager manager, PpChangeCase caseType)
        {
            return TextHelper.ApplyToText(manager, tr => tr.ChangeCase(caseType));
        }
    }

    [FeatureMetadata(Id = "BtnCharacterSpacing", Name = "Character Spacing", Tooltip = "Open PowerPoint's character-spacing gallery", IconData = "M8 4V20H10V4H8M14 4V20H16V4H14M4 10V14H6V10H4M18 10V14H20V10H18Z", Color = "#EAB308", Description = "Opens PowerPoint's spacing gallery or the custom spacing dialog for selected text.", DetailedHelpText = "### Character Spacing\n\nThe primary action opens PowerPoint's native Character Spacing gallery, whose presets (Very Tight, Tight, Normal, Loose, Very Loose) apply to the selection. Right-click **More Spacing...** to open PowerPoint's native Font dialog for exact values. PowerPoint's Font object model exposes no spacing property, so oPenEfficiency cannot apply exact presets directly: the native gallery and dialog are the supported contract.", MinSelection = 1)]
    public static class CharacterSpacingFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return TextHelper.ExecuteNativeCommand(manager, "CharacterSpacingGallery", "CharacterSpacingFeature.Execute");
        }

        
        public static bool ExecuteMoreSpacing(PowerPointManager manager)
        {
            return TextHelper.ExecuteNativeCommand(manager, "FontDialog", "CharacterSpacingFeature.ExecuteMoreSpacing");
        }
    }

    [FeatureMetadata(Id = "BtnAlignText", Name = "Align Text", Tooltip = "Open PowerPoint's text-alignment gallery", IconData = "M3 3h18v2H3V3m0 8h18v2H3v-2m0 8h18v2H3v-2z", Color = "#EAB308", Description = "Opens PowerPoint's text-alignment gallery or directly aligns text inside selected shapes.", DetailedHelpText = "### Align Text\n\nThe primary action opens PowerPoint's native text-alignment gallery. Right-click to set horizontal alignment, vertical anchoring, or one of the centered combinations directly. Direct actions affect the active text box or every compatible selected shape, skip shapes without a text frame, and create one PowerPoint Undo entry.", MinSelection = 1)]
    public static class AlignTextFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return TextHelper.ExecuteNativeCommand(manager, "AlignTextGallery", "AlignTextFeature.Execute");
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
            }, "AlignTextFeature.ExecuteSpecific");
        }
    }

    [FeatureMetadata(Id = "BtnTextDirection", Name = "Text Direction", Tooltip = "Open PowerPoint's text-direction gallery", IconData = "M9 3v13.59l-3.3-3.3L4.29 14.71L10 20.41l5.71-5.7l-1.41-1.42L11 16.59V3H9m7 4v2h5V7h-5m0 4v2h5v-2h-5z", Color = "#EAB308", Description = "Opens PowerPoint's text-direction gallery or directly changes text orientation.", DetailedHelpText = "### Text Direction\n\nThe primary action opens PowerPoint's native Text Direction gallery. Right-click to apply Horizontal, Rotate all text 90°, Rotate all text 270°, or Stacked to the active text box or every compatible selected shape. Direct actions create one PowerPoint Undo entry.", MinSelection = 1)]
    public static class TextDirectionFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return TextHelper.ExecuteNativeCommand(manager, "TextDirectionGallery", "TextDirectionFeature.Execute");
        }

        public static bool ExecuteSpecific(PowerPointManager manager, Office.MsoTextOrientation orientation)
        {
            return TextHelper.ApplyToTextFrame(
                manager,
                tf => tf.Orientation = orientation,
                "TextDirectionFeature.ExecuteSpecific");
        }
    }

    [FeatureMetadata(
        Id = "BtnNewSlide",
        Name = "New Slide",
        Tooltip = "Open PowerPoint's New Slide gallery",
        IconData = "M14 2H6c-1.1 0-1.99.9-1.99 2L4 20c0 1.1.89 2 1.99 2H18c1.1 0 2-.9 2-2V8l-6-6zm2 14h-3v3h-2v-3H8v-2h3v-3h2v3h3v2zm-3-7V3.5L18.5 9H13z",
        Color = "#10B981",
        Description = "Opens PowerPoint's native New Slide gallery without inserting a slide until a layout is chosen.",
        DetailedHelpText = "### New Slide\n\nOpen PowerPoint's native New Slide gallery and choose a layout from the available slide masters. Right-click to open the gallery, insert a blank slide after the current slide, or duplicate the current slide. Creation and Undo behavior remain native to PowerPoint where applicable.",
        MinSelection = 0,
        RequiredType = OfficeSelectionType.None)]
    public static class NewSlideFeature
    {
        public static bool Execute(PowerPointManager manager)
        {
            return TextHelper.ExecuteNativeCommand(
                manager,
                "SlideNewGallery",
                "NewSlideFeature.Execute");
        }

        public static bool InsertBlankAfterCurrent(PowerPointManager manager)
        {
            try
            {
                var app = manager?.GetApplication();
                var presentation = app?.ActivePresentation;
                var current = app?.ActiveWindow?.View?.Slide as Slide;
                if (presentation == null || current == null) return false;
                try { app.StartNewUndoEntry(); } catch { /* Slide insertion remains valid when undo grouping is unavailable. */ }
                presentation.Slides.Add(current.SlideIndex + 1, PpSlideLayout.ppLayoutBlank).Select();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "NewSlideFeature.InsertBlankAfterCurrent");
                return false;
            }
        }

        public static bool DuplicateCurrent(PowerPointManager manager)
        {
            try
            {
                var app = manager?.GetApplication();
                var current = app?.ActiveWindow?.View?.Slide as Slide;
                if (current == null) return false;
                try { app.StartNewUndoEntry(); } catch { /* Slide duplication remains valid when undo grouping is unavailable. */ }
                current.Duplicate().Select();
                return true;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "NewSlideFeature.DuplicateCurrent");
                return false;
            }
        }
    }
}
