using System;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Services.Attributes;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Features.Text
{
    [FeatureMetadata(
        Id = "BtnTextFormatPainter",
        Name = "Text Format Painter",
        Tooltip = "Copy text formatting from the first selected shape",
        IconData = "M4,3H16V7H18V11H16V21H8V11H6V7H4V3M6,5V7H14V5H6M10,11V19H14V11H10Z",
        Color = "#EAB308",
        Description = "Copies basic text and text-frame formatting from the first selected shape to the remaining selected shapes without changing their text.",
        DetailedHelpText = "### Text Format Painter\n\nSelect two or more text-bearing shapes. The first selected shape is the source; every other selected shape is a target. The primary action copies the source's first-character typography, paragraph alignment, vertical alignment, orientation, and text margins while retaining target text. Theme-linked text colors remain linked when PowerPoint exposes their theme value. Right-click **Apply Typography Only** to copy font family, size, color, and emphasis; **Apply Paragraphs & Bullets Only** to copy paragraph spacing, alignment, standard character or numbered bullets, and ruler indents by matching list level; or **Apply While Preserving Emphasis** to copy the basic font family, size, paragraph alignment, and text-frame settings without replacing each target's color, bold, italic, underline, subscript, or superscript emphasis. Picture bullets are left unchanged. Shapes without a text frame are skipped, the source must contain text, and each operation creates one PowerPoint Undo entry.",
        MinSelection = 2,
        RequiredType = OfficeSelectionType.Shapes)]
    public static class TextFormatPainterFeature
    {
        private enum PaintMode
        {
            Basic,
            TypographyOnly,
            ParagraphsAndBulletsOnly,
            PreserveEmphasis
        }

        public static bool Execute(PowerPointManager manager)
        {
            return Execute(manager, PaintMode.Basic);
        }

        public static bool ExecuteTypographyOnly(PowerPointManager manager)
        {
            return Execute(manager, PaintMode.TypographyOnly);
        }

        public static bool ExecuteParagraphsAndBulletsOnly(PowerPointManager manager)
        {
            return Execute(manager, PaintMode.ParagraphsAndBulletsOnly);
        }

        public static bool ExecutePreservingEmphasis(PowerPointManager manager)
        {
            return Execute(manager, PaintMode.PreserveEmphasis);
        }

        private static bool Execute(PowerPointManager manager, PaintMode mode)
        {
            if (manager == null) return false;

            try
            {
                var shapes = manager.GetSelectedShapes();
                if (shapes == null || shapes.Count < 2) return false;

                var source = shapes[1];
                if (!HasSourceText(source)) return false;

                try { manager.GetApplication().StartNewUndoEntry(); }
                catch (Exception ex) { ExceptionLogger.Log(ex, "TextFormatPainterFeature.Execute.StartNewUndoEntry"); }

                bool painted = false;
                bool failed = false;
                for (int index = 2; index <= shapes.Count; index++)
                {
                    var target = shapes[index];
                    if (!HasTextFrame(target)) continue;

                    try
                    {
                        CopyFormatting(source, target, mode);
                        painted = true;
                    }
                    catch (Exception ex)
                    {
                        failed = true;
                        ExceptionLogger.Log(ex, "TextFormatPainterFeature.Execute.Target." + mode);
                    }
                }

                return painted && !failed;
            }
            catch (Exception ex)
            {
                ExceptionLogger.Log(ex, "TextFormatPainterFeature.Execute." + mode);
                return false;
            }
        }

        private static bool HasTextFrame(Shape shape)
        {
            return shape != null && shape.HasTextFrame == Office.MsoTriState.msoTrue;
        }

        private static bool HasSourceText(Shape shape)
        {
            return HasTextFrame(shape) && shape.TextFrame.HasText == Office.MsoTriState.msoTrue;
        }

        private static void CopyFormatting(Shape source, Shape target, PaintMode mode)
        {
            switch (mode)
            {
                case PaintMode.TypographyOnly:
                    CopyTypography(source, target, true);
                    break;
                case PaintMode.ParagraphsAndBulletsOnly:
                    CopyParagraphsAndBullets(source, target);
                    break;
                case PaintMode.PreserveEmphasis:
                    CopyTypography(source, target, false);
                    CopyBasicLayout(source, target);
                    break;
                default:
                    CopyTypography(source, target, true);
                    CopyBasicLayout(source, target);
                    break;
            }
        }

        private static void CopyTypography(Shape source, Shape target, bool includeEmphasis)
        {
            var sourceFont = source.TextFrame.TextRange.Characters(1, 1).Font;
            var targetFont = target.TextFrame.TextRange.Font;

            targetFont.Name = sourceFont.Name;
            targetFont.NameAscii = sourceFont.NameAscii;
            targetFont.NameComplexScript = sourceFont.NameComplexScript;
            targetFont.NameFarEast = sourceFont.NameFarEast;
            targetFont.NameOther = sourceFont.NameOther;
            targetFont.Size = sourceFont.Size;

            if (!includeEmphasis) return;

            CopyColor(sourceFont.Color, targetFont.Color);
            targetFont.Bold = sourceFont.Bold;
            targetFont.Italic = sourceFont.Italic;
            targetFont.Underline = sourceFont.Underline;
            targetFont.Subscript = sourceFont.Subscript;
            targetFont.Superscript = sourceFont.Superscript;
        }

        private static void CopyBasicLayout(Shape source, Shape target)
        {
            var sourceText = source.TextFrame.TextRange;
            var targetText = target.TextFrame.TextRange;

            targetText.ParagraphFormat.Alignment = sourceText.Paragraphs(1, 1).ParagraphFormat.Alignment;
            target.TextFrame.VerticalAnchor = source.TextFrame.VerticalAnchor;
            target.TextFrame.Orientation = source.TextFrame.Orientation;
            target.TextFrame.MarginLeft = source.TextFrame.MarginLeft;
            target.TextFrame.MarginRight = source.TextFrame.MarginRight;
            target.TextFrame.MarginTop = source.TextFrame.MarginTop;
            target.TextFrame.MarginBottom = source.TextFrame.MarginBottom;
        }

        private static void CopyParagraphsAndBullets(Shape source, Shape target)
        {
            var sourceText = source.TextFrame.TextRange;
            var targetText = target.TextFrame.TextRange;
            int targetParagraphCount = Math.Max(1, targetText.Paragraphs().Count);

            for (int index = 1; index <= targetParagraphCount; index++)
            {
                var targetParagraph = targetText.Paragraphs(index, 1);
                var sourceParagraph = FindSourceParagraphForLevel(sourceText, targetParagraph.IndentLevel);
                CopyParagraphFormat(sourceParagraph.ParagraphFormat, targetParagraph.ParagraphFormat);
                CopyBulletFormat(sourceParagraph.ParagraphFormat.Bullet, targetParagraph.ParagraphFormat.Bullet);
            }

            CopyRulerLevels(source.TextFrame.Ruler.Levels, target.TextFrame.Ruler.Levels);
        }

        private static TextRange FindSourceParagraphForLevel(TextRange sourceText, int indentLevel)
        {
            int sourceParagraphCount = Math.Max(1, sourceText.Paragraphs().Count);
            for (int index = 1; index <= sourceParagraphCount; index++)
            {
                var paragraph = sourceText.Paragraphs(index, 1);
                if (paragraph.IndentLevel == indentLevel) return paragraph;
            }

            return sourceText.Paragraphs(1, 1);
        }

        private static void CopyParagraphFormat(ParagraphFormat source, ParagraphFormat target)
        {
            target.Alignment = source.Alignment;
            target.BaseLineAlignment = source.BaseLineAlignment;
            target.FarEastLineBreakControl = source.FarEastLineBreakControl;
            target.HangingPunctuation = source.HangingPunctuation;
            target.LineRuleAfter = source.LineRuleAfter;
            target.LineRuleBefore = source.LineRuleBefore;
            target.LineRuleWithin = source.LineRuleWithin;
            target.SpaceAfter = source.SpaceAfter;
            target.SpaceBefore = source.SpaceBefore;
            target.SpaceWithin = source.SpaceWithin;
            target.TextDirection = source.TextDirection;
            target.WordWrap = source.WordWrap;
        }

        private static void CopyBulletFormat(BulletFormat source, BulletFormat target)
        {
            if (source.Type == PpBulletType.ppBulletPicture ||
                source.Type == PpBulletType.ppBulletMixed)
            {
                return;
            }

            target.Visible = source.Visible;
            target.Type = source.Type;
            if (source.Type == PpBulletType.ppBulletNone) return;

            target.RelativeSize = source.RelativeSize;
            target.UseTextColor = source.UseTextColor;
            target.UseTextFont = source.UseTextFont;

            if (source.UseTextFont != Office.MsoTriState.msoTrue)
            {
                target.Font.Name = source.Font.Name;
            }
            if (source.UseTextColor != Office.MsoTriState.msoTrue)
            {
                CopyColor(source.Font.Color, target.Font.Color);
            }

            if (source.Type == PpBulletType.ppBulletUnnumbered)
            {
                target.Character = source.Character;
            }
            else if (source.Type == PpBulletType.ppBulletNumbered)
            {
                target.Style = source.Style;
                target.StartValue = source.StartValue;
            }
        }

        private static void CopyRulerLevels(RulerLevels source, RulerLevels target)
        {
            int count = Math.Min(source.Count, target.Count);
            for (int index = 1; index <= count; index++)
            {
                target[index].FirstMargin = source[index].FirstMargin;
                target[index].LeftMargin = source[index].LeftMargin;
            }
        }

        private static void CopyColor(ColorFormat source, ColorFormat target)
        {
            if (source.Type == Office.MsoColorType.msoColorTypeScheme &&
                source.ObjectThemeColor != Office.MsoThemeColorIndex.msoNotThemeColor &&
                source.ObjectThemeColor != Office.MsoThemeColorIndex.msoThemeColorMixed)
            {
                target.ObjectThemeColor = source.ObjectThemeColor;
                target.TintAndShade = source.TintAndShade;
                target.Brightness = source.Brightness;
                return;
            }

            target.RGB = source.RGB;
        }
    }
}
