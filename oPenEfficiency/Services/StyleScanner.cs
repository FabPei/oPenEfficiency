using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Services
{
    public class StyleError
    {
        public string RuleId { get; set; }
        public string ErrorType { get; set; } // "Font", "Font Size", "Fill Color", "Line Color", "Text Color", "Distorted Image", "Overlapping Text", "Protection Area"
        public string Value { get; set; }
        public int ColorValue { get; set; }
        public int SlideNumber { get; set; }
        public string ShapeName { get; set; }
        public string SuggestedCorrectionValue { get; set; }
        public int TextStart { get; set; } = -1;
        public int TextLength { get; set; } = 0;
    }

    public static class StyleCheckRuleIgnore
    {
        private const string LegacyStyleTag = "oPE_IgnoreStyleCheck";
        private const string LegacyConsistencyTag = "oPE_IgnoreConsistencyCheck";

        public static bool IsIgnored(Shape shape, string category, string ruleId)
        {
            if (shape == null) return false;
            return HasTag(shape.Tags, GetLegacyTag(category)) ||
                   HasTag(shape.Tags, GetRuleTag(category, ruleId));
        }

        public static bool IsIgnored(Slide slide, string category, string ruleId)
        {
            if (slide == null) return false;
            return HasTag(slide.Tags, GetLegacyTag(category)) ||
                   HasTag(slide.Tags, GetRuleTag(category, ruleId));
        }

        public static void Ignore(Shape shape, string category, string ruleId)
        {
            if (shape == null) return;
            shape.Tags.Add(GetRuleTag(category, ruleId), "True");
        }

        public static void Ignore(Slide slide, string category, string ruleId)
        {
            if (slide == null) return;
            slide.Tags.Add(GetRuleTag(category, ruleId), "True");
        }

        private static string GetLegacyTag(string category)
        {
            return string.Equals(category, "Consistency", StringComparison.OrdinalIgnoreCase)
                ? LegacyConsistencyTag
                : LegacyStyleTag;
        }

        private static string GetRuleTag(string category, string ruleId)
        {
            string prefix = string.Equals(category, "Consistency", StringComparison.OrdinalIgnoreCase)
                ? "oPE_SC_C_"
                : "oPE_SC_S_";
            string safeRule = new string((ruleId ?? "Unknown")
                .Where(char.IsLetterOrDigit)
                .ToArray());
            return prefix + safeRule;
        }

        private static bool HasTag(Tags tags, string name)
        {
            try
            {
                return string.Equals(tags[name], "True", StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }
    }

    public class StyleScanner
    {
        public static List<StyleError> ScanSlides(Application app, List<Slide> slides)
        {
            var errors = new List<StyleError>();
            if (slides == null || slides.Count == 0) return errors;

            var pres = app.ActivePresentation;
            var themeFonts = GetThemeFonts(pres);
            
            var config = JsonService.Load<Models.SidebarConfig>("sidebar_config.json");
            var correctColors = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var correctFonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var correctFontSizes = new HashSet<float>();
            
            float[] protectionArea = null;
            float? targetLineSpacing = null;
            float? targetSpaceBefore = null;
            float? targetSpaceAfter = null;

            if (config != null)
            {
                if (!string.IsNullOrWhiteSpace(config.StyleCheckCorrectColors))
                {
                    foreach (var c in config.StyleCheckCorrectColors.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        var trimmed = c.Trim();
                        if (!trimmed.StartsWith("#")) trimmed = "#" + trimmed;
                        correctColors.Add(trimmed);
                    }
                }
                
                if (!string.IsNullOrWhiteSpace(config.StyleCheckCorrectFonts))
                {
                    foreach (var f in config.StyleCheckCorrectFonts.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        correctFonts.Add(f.Trim());
                    }
                }

                if (!string.IsNullOrWhiteSpace(config.StyleCheckCorrectFontSizes))
                {
                    foreach (var s in config.StyleCheckCorrectFontSizes.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        if (float.TryParse(s.Trim(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out float size))
                        {
                            correctFontSizes.Add(size);
                        }
                    }
                }

                if (!string.IsNullOrWhiteSpace(config.StyleCheckLineSpacing) && float.TryParse(config.StyleCheckLineSpacing, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out float ls))
                    targetLineSpacing = ls;
                if (!string.IsNullOrWhiteSpace(config.StyleCheckSpaceBefore) && float.TryParse(config.StyleCheckSpaceBefore, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out float sb))
                    targetSpaceBefore = sb;
                if (!string.IsNullOrWhiteSpace(config.StyleCheckSpaceAfter) && float.TryParse(config.StyleCheckSpaceAfter, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out float sa))
                    targetSpaceAfter = sa;

                if (!string.IsNullOrWhiteSpace(config.StyleCheckProtectionArea))
                {
                    var parts = config.StyleCheckProtectionArea.Split(',');
                    if (parts.Length == 4)
                    {
                        if (float.TryParse(parts[0], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out float x) && 
                            float.TryParse(parts[1], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out float y) &&
                            float.TryParse(parts[2], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out float w) && 
                            float.TryParse(parts[3], System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out float h))
                        {
                            protectionArea = new float[] { x, y, w, h };
                        }
                    }
                }
            }

            foreach (var slide in slides)
            {
                var textShapes = new List<Shape>();

                foreach (Shape shape in slide.Shapes)
                {
                    if (shape.Type != Office.MsoShapeType.msoGroup && shape.HasTextFrame == Office.MsoTriState.msoTrue)
                    {
                        textShapes.Add(shape);
                    }
                    ScanShape(shape, slide.SlideNumber, themeFonts, correctFonts, correctColors, correctFontSizes, protectionArea, targetLineSpacing, targetSpaceBefore, targetSpaceAfter, errors);
                }

                // Phase 3: Spatial overlap detection for text
                for (int i = 0; i < textShapes.Count; i++)
                {
                    for (int j = i + 1; j < textShapes.Count; j++)
                    {
                        try {
                            if (!StyleCheckRuleIgnore.IsIgnored(textShapes[i], "Style", "OverlappingText") &&
                                !StyleCheckRuleIgnore.IsIgnored(textShapes[j], "Style", "OverlappingText") &&
                                IsOverlapping(textShapes[i], textShapes[j]))
                            {
                                errors.Add(new StyleError {
                                    RuleId = "OverlappingText",
                                    ErrorType = "Overlapping Text",
                                    Value = $"Shapes overlap",
                                    SlideNumber = slide.SlideNumber,
                                    ShapeName = textShapes[i].Name
                                });
                            }
                        }
                        catch (Exception ex) { ReportOptionalScanFailure(ex, "StyleScanner.Scan.OverlappingText"); }
                    }
                }
            }

            return errors;
        }

        private static bool IsOverlapping(Shape a, Shape b)
        {
            return a.Left < b.Left + b.Width &&
                   a.Left + a.Width > b.Left &&
                   a.Top < b.Top + b.Height &&
                   a.Top + a.Height > b.Top;
        }

        private static HashSet<string> GetThemeFonts(Presentation pres)
        {
            var fonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                var theme = pres.SlideMaster.Theme;
                var fontScheme = theme.ThemeFontScheme;

                var langs = new[] { 
                    Office.MsoFontLanguageIndex.msoThemeLatin, 
                    Office.MsoFontLanguageIndex.msoThemeComplexScript, 
                    Office.MsoFontLanguageIndex.msoThemeEastAsian 
                };

                foreach (var lang in langs)
                {
                    try { fonts.Add(fontScheme.MajorFont.Item(lang).Name); }
                    catch (Exception ex) { ReportOptionalScanFailure(ex, "StyleScanner.GetThemeFonts.MajorFont"); }
                    try { fonts.Add(fontScheme.MinorFont.Item(lang).Name); }
                    catch (Exception ex) { ReportOptionalScanFailure(ex, "StyleScanner.GetThemeFonts.MinorFont"); }
                }
            }
            catch (Exception ex) { ReportOptionalScanFailure(ex, "StyleScanner.GetThemeFonts"); }
            return fonts;
        }

        private static void ScanShape(Shape shape, int slideNum, HashSet<string> themeFonts, HashSet<string> correctFonts, HashSet<string> correctColors, HashSet<float> correctFontSizes, float[] protectionArea, float? targetLineSpacing, float? targetSpaceBefore, float? targetSpaceAfter, List<StyleError> errors)
        {
            try
            {
                if (StyleCheckRuleIgnore.IsIgnored(shape, "Style", "All")) return;

                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    foreach (Shape child in shape.GroupItems) ScanShape(child, slideNum, themeFonts, correctFonts, correctColors, correctFontSizes, protectionArea, targetLineSpacing, targetSpaceBefore, targetSpaceAfter, errors);
                    return;
                }
            }
            catch (Exception ex) { ReportOptionalScanFailure(ex, "StyleScanner.ScanShape.Group"); }

            // Phase 3: Layout Protection Area Check
            if (protectionArea != null && !StyleCheckRuleIgnore.IsIgnored(shape, "Style", "ProtectionArea"))
            {
                try
                {
                    if (shape.Left < protectionArea[0] + protectionArea[2] &&
                        shape.Left + shape.Width > protectionArea[0] &&
                        shape.Top < protectionArea[1] + protectionArea[3] &&
                        shape.Top + shape.Height > protectionArea[1])
                    {
                        // Check if it's not a placeholder (placeholders can be in protection area depending on rules, but let's flag if it is an overlapping shape)
                        if (shape.Type != Office.MsoShapeType.msoPlaceholder)
                        {
                            errors.Add(new StyleError { RuleId = "ProtectionArea", ErrorType = "Protection Area", Value = "Violates protection zone", SlideNumber = slideNum, ShapeName = shape.Name });
                        }
                    }
                }
                catch (Exception ex) { ReportOptionalScanFailure(ex, "StyleScanner.ScanShape.ProtectionArea"); }
            }

            // Phase 3: Distorted Images
            try
            {
                if (shape.Type == Office.MsoShapeType.msoPicture &&
                    !StyleCheckRuleIgnore.IsIgnored(shape, "Style", "DistortedImage"))
                {
                    float currentWidth = shape.Width;
                    float currentHeight = shape.Height;
                    
                    Shape duplicate = shape.Duplicate()[1];
                    duplicate.ScaleWidth(1f, Office.MsoTriState.msoTrue);
                    duplicate.ScaleHeight(1f, Office.MsoTriState.msoTrue);
                    
                    float intrinsicWidth = duplicate.Width;
                    float intrinsicHeight = duplicate.Height;
                    
                    duplicate.Delete();
                    
                    if (intrinsicWidth > 0 && intrinsicHeight > 0 && currentHeight > 0)
                    {
                        float originalRatio = intrinsicWidth / intrinsicHeight;
                        float currentRatio = currentWidth / currentHeight;
                        
                        if (Math.Abs(originalRatio - currentRatio) > 0.05f) // 5% variance
                        {
                            errors.Add(new StyleError { RuleId = "DistortedImage", ErrorType = "Distorted Image", Value = "Aspect ratio skewed", SlideNumber = slideNum, ShapeName = shape.Name });
                        }
                    }
                }
            }
            catch (Exception ex) { ReportOptionalScanFailure(ex, "StyleScanner.ScanShape.DistortedImage"); }

            // Table
            try
            {
                if (shape.HasTable == Office.MsoTriState.msoTrue)
                {
                    for (int r = 1; r <= shape.Table.Rows.Count; r++)
                    {
                        for (int c = 1; c <= shape.Table.Columns.Count; c++)
                        {
                            var cell = shape.Table.Cell(r, c);
                            try
                            {
                                if (!StyleCheckRuleIgnore.IsIgnored(shape, "Style", "FillColor") &&
                                    cell.Shape.Fill.Visible == Office.MsoTriState.msoTrue &&
                                    cell.Shape.Fill.ForeColor.ObjectThemeColor == Office.MsoThemeColorIndex.msoNotThemeColor)
                                {
                                    AddColorError(errors, "FillColor", "Fill Color", cell.Shape.Fill.ForeColor.RGB, slideNum, shape.Name, correctColors, allowCorrection: false);
                                }
                            }
                            catch (Exception ex) { ReportOptionalScanFailure(ex, "StyleScanner.ScanShape.TableFill"); }

                            try
                            {
                                if (cell.Shape.HasTextFrame == Office.MsoTriState.msoTrue && cell.Shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                                {
                                    ScanTextRange(cell.Shape.TextFrame.TextRange, shape, slideNum, shape.Name, themeFonts, correctFonts, correctColors, correctFontSizes, targetLineSpacing, targetSpaceBefore, targetSpaceAfter, errors);
                                }
                            }
                            catch (Exception ex) { ReportOptionalScanFailure(ex, "StyleScanner.ScanShape.TableText"); }
                        }
                    }
                    return;
                }
            }
            catch (Exception ex) { ReportOptionalScanFailure(ex, "StyleScanner.ScanShape.Table"); }

            // Fill Color
            try
            {
                if (!StyleCheckRuleIgnore.IsIgnored(shape, "Style", "FillColor") &&
                    shape.Fill.Visible == Office.MsoTriState.msoTrue && shape.Type != Office.MsoShapeType.msoLine)
                {
                    if (shape.Fill.Type == Office.MsoFillType.msoFillSolid && 
                        shape.Fill.ForeColor.ObjectThemeColor == Office.MsoThemeColorIndex.msoNotThemeColor)
                    {
                        AddColorError(errors, "FillColor", "Fill Color", shape.Fill.ForeColor.RGB, slideNum, shape.Name, correctColors);
                    }
                }
            }
            catch (Exception ex) { ReportOptionalScanFailure(ex, "StyleScanner.ScanShape.FillColor"); }

            // Line Color
            try
            {
                if (!StyleCheckRuleIgnore.IsIgnored(shape, "Style", "LineColor") &&
                    shape.Line.Visible == Office.MsoTriState.msoTrue && shape.Line.Weight > 0)
                {
                    if (shape.Line.ForeColor.ObjectThemeColor == Office.MsoThemeColorIndex.msoNotThemeColor)
                    {
                        AddColorError(errors, "LineColor", "Line Color", shape.Line.ForeColor.RGB, slideNum, shape.Name, correctColors);
                    }
                }
            }
            catch (Exception ex) { ReportOptionalScanFailure(ex, "StyleScanner.ScanShape.LineColor"); }

            // Text (Fonts)
            try
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    ScanTextOverflow(shape, slideNum, errors);
                    ScanTextRange(shape.TextFrame.TextRange, shape, slideNum, shape.Name, themeFonts, correctFonts, correctColors, correctFontSizes, targetLineSpacing, targetSpaceBefore, targetSpaceAfter, errors);
                }
            }
            catch (Exception ex) { ReportOptionalScanFailure(ex, "StyleScanner.ScanShape.Text"); }
        }

        private static void ScanTextRange(TextRange textRange, Shape ignoreOwner, int slideNum, string shapeName, HashSet<string> themeFonts, HashSet<string> correctFonts, HashSet<string> correctColors, HashSet<float> correctFontSizes, float? targetLineSpacing, float? targetSpaceBefore, float? targetSpaceAfter, List<StyleError> errors)
        {
            try
            {
                try
            {
                int paraCount = 0;
                try { paraCount = textRange.Paragraphs().Count; }
                catch (Exception ex) { ReportOptionalScanFailure(ex, "StyleScanner.ScanTextRange.ParagraphCount"); }

                for (int i = 1; i <= paraCount; i++)
                {
                    TextRange para = null;
                    try { para = textRange.Paragraphs(i); }
                    catch (Exception ex) { ReportOptionalScanFailure(ex, "StyleScanner.ScanTextRange.Paragraph"); }

                    if (para == null) continue;

                    var pf = para.ParagraphFormat;
                    int pStart = para.Start;
                    int pLength = para.Length;

                    if (!StyleCheckRuleIgnore.IsIgnored(ignoreOwner, "Style", "ParagraphSpacing") &&
                        targetLineSpacing.HasValue && Math.Abs(pf.SpaceWithin - targetLineSpacing.Value) > 0.1f)
                    {
                        errors.Add(new StyleError { RuleId = "ParagraphSpacing", ErrorType = "Paragraph Spacing", Value = $"Line Spacing: {pf.SpaceWithin}", SlideNumber = slideNum, ShapeName = shapeName, TextStart = pStart, TextLength = pLength });
                    }
                    if (!StyleCheckRuleIgnore.IsIgnored(ignoreOwner, "Style", "ParagraphSpacing") &&
                        targetSpaceBefore.HasValue && Math.Abs(pf.SpaceBefore - targetSpaceBefore.Value) > 0.1f)
                    {
                        errors.Add(new StyleError { RuleId = "ParagraphSpacing", ErrorType = "Paragraph Spacing", Value = $"Space Before: {pf.SpaceBefore}", SlideNumber = slideNum, ShapeName = shapeName, TextStart = pStart, TextLength = pLength });
                    }
                    if (!StyleCheckRuleIgnore.IsIgnored(ignoreOwner, "Style", "ParagraphSpacing") &&
                        targetSpaceAfter.HasValue && Math.Abs(pf.SpaceAfter - targetSpaceAfter.Value) > 0.1f)
                    {
                        errors.Add(new StyleError { RuleId = "ParagraphSpacing", ErrorType = "Paragraph Spacing", Value = $"Space After: {pf.SpaceAfter}", SlideNumber = slideNum, ShapeName = shapeName, TextStart = pStart, TextLength = pLength });
                    }
                }
            }
            catch (Exception ex) { ReportOptionalScanFailure(ex, "StyleScanner.ScanTextRange.Paragraphs"); }
            
            int runCount = 0;
                try { runCount = textRange.Runs().Count; }
                catch (Exception ex) { ReportOptionalScanFailure(ex, "StyleScanner.ScanTextRange.RunCount"); }

                int iterations = Math.Max(1, runCount);

                for (int i = 1; i <= iterations; i++)
                {
                    TextRange run = null;
                    try { run = runCount == 0 ? textRange : textRange.Runs(i); }
                    catch (Exception ex) { ReportOptionalScanFailure(ex, "StyleScanner.ScanTextRange.Run"); }

                    if (run == null) continue;

                    int tStart = run.Start;
                    int tLength = run.Length;

                    var fontName = run.Font.Name;
                    if (!StyleCheckRuleIgnore.IsIgnored(ignoreOwner, "Style", "Font") &&
                        !string.IsNullOrEmpty(fontName) && !fontName.StartsWith("+") && !themeFonts.Contains(fontName) && !correctFonts.Contains(fontName))
                    {
                        errors.Add(new StyleError
                        {
                            RuleId = "Font",
                            ErrorType = "Font",
                            Value = fontName,
                            SlideNumber = slideNum,
                            ShapeName = shapeName,
                            TextStart = tStart,
                            TextLength = tLength
                        });
                    }

                    // Font Size Validation
                    if (!StyleCheckRuleIgnore.IsIgnored(ignoreOwner, "Style", "FontSize") &&
                        correctFontSizes.Count > 0 && !correctFontSizes.Contains(run.Font.Size))
                    {
                        errors.Add(new StyleError
                        {
                            RuleId = "FontSize",
                            ErrorType = "Font Size",
                            Value = run.Font.Size.ToString(),
                            SlideNumber = slideNum,
                            ShapeName = shapeName,
                            TextStart = tStart,
                            TextLength = tLength
                        });
                    }

                    try
                    {
                        if (!StyleCheckRuleIgnore.IsIgnored(ignoreOwner, "Style", "TextColor") &&
                            run.Font.Color.ObjectThemeColor == Office.MsoThemeColorIndex.msoNotThemeColor)
                        {
                            AddColorError(errors, "TextColor", "Text Color", run.Font.Color.RGB, slideNum, shapeName, correctColors, tStart, tLength);
                        }
                    }
                    catch (Exception ex) { ReportOptionalScanFailure(ex, "StyleScanner.ScanTextRange.TextColor"); }
                }
            }
            catch (Exception ex) { ReportOptionalScanFailure(ex, "StyleScanner.ScanTextRange"); }
        }

        private static void ScanTextOverflow(Shape shape, int slideNum, List<StyleError> errors)
        {
            if (StyleCheckRuleIgnore.IsIgnored(shape, "Style", "TextOverflow")) return;

            try
            {
                if (shape.TextFrame.AutoSize == PpAutoSize.ppAutoSizeShapeToFitText) return;

                float availableHeight = shape.Height -
                    shape.TextFrame.MarginTop -
                    shape.TextFrame.MarginBottom;
                float requiredHeight = shape.TextFrame.TextRange.BoundHeight;

                if (availableHeight > 0 && requiredHeight > availableHeight + 1f)
                {
                    errors.Add(new StyleError
                    {
                        RuleId = "TextOverflow",
                        ErrorType = "Text Overflow",
                        Value = "Text exceeds the available shape height",
                        SlideNumber = slideNum,
                        ShapeName = shape.Name
                    });
                }
            }
            catch (Exception ex) { ReportOptionalScanFailure(ex, "StyleScanner.ScanTextOverflow"); }
        }

        private static void ReportOptionalScanFailure(Exception exception, string context)
        {
            if (!ExceptionPolicy.IsExpectedComFailure(exception))
            {
                ExceptionLogger.Log(exception, context);
            }
            // A single inaccessible COM property must not prevent scanning the remaining shapes.
        }

        private static void AddColorError(List<StyleError> errors, string ruleId, string type, int rgb, int slideNum, string shapeName, HashSet<string> correctColors, int start = -1, int length = 0, bool allowCorrection = true)
        {
            string hex = $"#{((rgb & 0xFF)):X2}{((rgb & 0xFF00) >> 8):X2}{((rgb & 0xFF0000) >> 16):X2}";
            
            if (correctColors != null && correctColors.Contains(hex)) return;
            
            string suggestedCorrection = null;
            if (allowCorrection && correctColors != null && correctColors.Count > 0)
            {
                suggestedCorrection = ColorMath.FindClosestColor(rgb, correctColors);
            }

            errors.Add(new StyleError
            {
                RuleId = ruleId,
                ErrorType = type,
                Value = hex,
                ColorValue = rgb,
                SlideNumber = slideNum,
                ShapeName = shapeName,
                SuggestedCorrectionValue = suggestedCorrection,
                TextStart = start,
                TextLength = length
            });
        }
    }
}
