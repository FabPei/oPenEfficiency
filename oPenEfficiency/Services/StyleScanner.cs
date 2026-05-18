using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using oPenEfficiency.Utils;

namespace oPenEfficiency.Services
{
    public class StyleError
    {
        public string ErrorType { get; set; } // "Font", "Font Size", "Fill Color", "Line Color", "Text Color", "Distorted Image", "Overlapping Text", "Protection Area"
        public string Value { get; set; }
        public int ColorValue { get; set; }
        public int SlideNumber { get; set; }
        public string ShapeName { get; set; }
        public string SuggestedCorrectionValue { get; set; }
        public int TextStart { get; set; } = -1;
        public int TextLength { get; set; } = 0;
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
                            // Check ignore tags
                            bool ignoreI = false;
                            bool ignoreJ = false;
                            try { ignoreI = textShapes[i].Tags["oPE_IgnoreStyleCheck"] == "True"; } catch { }
                            try { ignoreJ = textShapes[j].Tags["oPE_IgnoreStyleCheck"] == "True"; } catch { }

                            if (!ignoreI && !ignoreJ && IsOverlapping(textShapes[i], textShapes[j]))
                            {
                                errors.Add(new StyleError {
                                    ErrorType = "Overlapping Text",
                                    Value = $"Shapes overlap",
                                    SlideNumber = slide.SlideNumber,
                                    ShapeName = textShapes[i].Name
                                });
                            }
                        } catch { }
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
                    try { fonts.Add(fontScheme.MajorFont.Item(lang).Name); } catch { }
                    try { fonts.Add(fontScheme.MinorFont.Item(lang).Name); } catch { }
                }
            }
            catch { }
            return fonts;
        }

        private static void ScanShape(Shape shape, int slideNum, HashSet<string> themeFonts, HashSet<string> correctFonts, HashSet<string> correctColors, HashSet<float> correctFontSizes, float[] protectionArea, float? targetLineSpacing, float? targetSpaceBefore, float? targetSpaceAfter, List<StyleError> errors)
        {
            try
            {
                // Phase 5: The "Ignore" System
                try {
                    if (shape.Tags["oPE_IgnoreStyleCheck"] == "True") return;
                } catch { }

                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    foreach (Shape child in shape.GroupItems) ScanShape(child, slideNum, themeFonts, correctFonts, correctColors, correctFontSizes, protectionArea, targetLineSpacing, targetSpaceBefore, targetSpaceAfter, errors);
                    return;
                }
            }
            catch { }

            // Phase 3: Layout Protection Area Check
            if (protectionArea != null)
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
                            errors.Add(new StyleError { ErrorType = "Protection Area", Value = "Violates protection zone", SlideNumber = slideNum, ShapeName = shape.Name });
                        }
                    }
                }
                catch { }
            }

            // Phase 3: Distorted Images
            try
            {
                if (shape.Type == Office.MsoShapeType.msoPicture)
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
                            errors.Add(new StyleError { ErrorType = "Distorted Image", Value = "Aspect ratio skewed", SlideNumber = slideNum, ShapeName = shape.Name });
                        }
                    }
                }
            }
            catch { }

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
                                if (cell.Shape.Fill.Visible == Office.MsoTriState.msoTrue && 
                                    cell.Shape.Fill.ForeColor.ObjectThemeColor == Office.MsoThemeColorIndex.msoNotThemeColor)
                                {
                                    AddColorError(errors, "Fill Color", cell.Shape.Fill.ForeColor.RGB, slideNum, shape.Name, correctColors);
                                }
                            }
                            catch { }

                            try
                            {
                                if (cell.Shape.HasTextFrame == Office.MsoTriState.msoTrue && cell.Shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                                {
                                    ScanTextRange(cell.Shape.TextFrame.TextRange, slideNum, shape.Name, themeFonts, correctFonts, correctColors, correctFontSizes, targetLineSpacing, targetSpaceBefore, targetSpaceAfter, errors);
                                }
                            }
                            catch { }
                        }
                    }
                    return;
                }
            }
            catch { }

            // Fill Color
            try
            {
                if (shape.Fill.Visible == Office.MsoTriState.msoTrue && shape.Type != Office.MsoShapeType.msoLine)
                {
                    if (shape.Fill.Type == Office.MsoFillType.msoFillSolid && 
                        shape.Fill.ForeColor.ObjectThemeColor == Office.MsoThemeColorIndex.msoNotThemeColor)
                    {
                        AddColorError(errors, "Fill Color", shape.Fill.ForeColor.RGB, slideNum, shape.Name, correctColors);
                    }
                }
            }
            catch { }

            // Line Color
            try
            {
                if (shape.Line.Visible == Office.MsoTriState.msoTrue && shape.Line.Weight > 0)
                {
                    if (shape.Line.ForeColor.ObjectThemeColor == Office.MsoThemeColorIndex.msoNotThemeColor)
                    {
                        AddColorError(errors, "Line Color", shape.Line.ForeColor.RGB, slideNum, shape.Name, correctColors);
                    }
                }
            }
            catch { }

            // Text (Fonts)
            try
            {
                if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                {
                    ScanTextRange(shape.TextFrame.TextRange, slideNum, shape.Name, themeFonts, correctFonts, correctColors, correctFontSizes, targetLineSpacing, targetSpaceBefore, targetSpaceAfter, errors);
                }
            }
            catch { }
        }

        private static void ScanTextRange(TextRange textRange, int slideNum, string shapeName, HashSet<string> themeFonts, HashSet<string> correctFonts, HashSet<string> correctColors, HashSet<float> correctFontSizes, float? targetLineSpacing, float? targetSpaceBefore, float? targetSpaceAfter, List<StyleError> errors)
        {
            try
            {
                try
            {
                int paraCount = 0;
                try { paraCount = textRange.Paragraphs().Count; } catch { }

                for (int i = 1; i <= paraCount; i++)
                {
                    TextRange para = null;
                    try { para = textRange.Paragraphs(i); } catch { }

                    if (para == null) continue;

                    var pf = para.ParagraphFormat;
                    int pStart = para.Start;
                    int pLength = para.Length;

                    if (targetLineSpacing.HasValue && Math.Abs(pf.SpaceWithin - targetLineSpacing.Value) > 0.1f)
                    {
                        errors.Add(new StyleError { ErrorType = "Paragraph Spacing", Value = $"Line Spacing: {pf.SpaceWithin}", SlideNumber = slideNum, ShapeName = shapeName, TextStart = pStart, TextLength = pLength });
                    }
                    if (targetSpaceBefore.HasValue && Math.Abs(pf.SpaceBefore - targetSpaceBefore.Value) > 0.1f)
                    {
                        errors.Add(new StyleError { ErrorType = "Paragraph Spacing", Value = $"Space Before: {pf.SpaceBefore}", SlideNumber = slideNum, ShapeName = shapeName, TextStart = pStart, TextLength = pLength });
                    }
                    if (targetSpaceAfter.HasValue && Math.Abs(pf.SpaceAfter - targetSpaceAfter.Value) > 0.1f)
                    {
                        errors.Add(new StyleError { ErrorType = "Paragraph Spacing", Value = $"Space After: {pf.SpaceAfter}", SlideNumber = slideNum, ShapeName = shapeName, TextStart = pStart, TextLength = pLength });
                    }
                }
            }
            catch { }
            
            int runCount = 0;
                try { runCount = textRange.Runs().Count; } catch { }

                int iterations = Math.Max(1, runCount);

                for (int i = 1; i <= iterations; i++)
                {
                    TextRange run = null;
                    try { run = runCount == 0 ? textRange : textRange.Runs(i); } catch { }

                    if (run == null) continue;

                    int tStart = run.Start;
                    int tLength = run.Length;

                    var fontName = run.Font.Name;
                    if (!string.IsNullOrEmpty(fontName) && !fontName.StartsWith("+") && !themeFonts.Contains(fontName) && !correctFonts.Contains(fontName))
                    {
                        errors.Add(new StyleError
                        {
                            ErrorType = "Font",
                            Value = fontName,
                            SlideNumber = slideNum,
                            ShapeName = shapeName,
                            TextStart = tStart,
                            TextLength = tLength
                        });
                    }

                    // Font Size Validation
                    if (correctFontSizes.Count > 0 && !correctFontSizes.Contains(run.Font.Size))
                    {
                        errors.Add(new StyleError
                        {
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
                        if (run.Font.Color.ObjectThemeColor == Office.MsoThemeColorIndex.msoNotThemeColor)
                        {
                            AddColorError(errors, "Text Color", run.Font.Color.RGB, slideNum, shapeName, correctColors, tStart, tLength);
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }

        private static void AddColorError(List<StyleError> errors, string type, int rgb, int slideNum, string shapeName, HashSet<string> correctColors, int start = -1, int length = 0)
        {
            string hex = $"#{((rgb & 0xFF)):X2}{((rgb & 0xFF00) >> 8):X2}{((rgb & 0xFF0000) >> 16):X2}";
            
            if (correctColors != null && correctColors.Contains(hex)) return;
            
            string suggestedCorrection = null;
            if (correctColors != null && correctColors.Count > 0)
            {
                suggestedCorrection = ColorMath.FindClosestColor(rgb, correctColors);
            }

            errors.Add(new StyleError
            {
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