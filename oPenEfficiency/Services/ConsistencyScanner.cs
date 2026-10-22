using System;
using oPenEfficiency.Utils;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.Services
{
    public class ConsistencyError
    {
        public string RuleId { get; set; }
        public string ErrorType { get; set; }
        public string Description { get; set; }
        public int SlideNumber { get; set; }
        public string ShapeName { get; set; }
        public string SuggestedCorrection { get; set; }
        public int TextStart { get; set; } = -1;
        public int TextLength { get; set; } = 0;
    }

    public class ConsistencyScanner
    {
        private const int RecommendedTitleLength = 120;

        public static List<ConsistencyError> ScanSlides(Application app, List<Slide> slides)
        {
            var errors = new List<ConsistencyError>();
            if (slides == null || slides.Count == 0) return errors;

            bool footerExpected = IsFooterExpected(app, slides);

            foreach (var slide in slides)
            {
                bool slideNumberFound = false;
                bool sourceFound = false;
                bool containsDataVisual = false;
                Shape titleShape = null;

                foreach (Shape shape in EnumerateShapes(slide))
                {
                    if (IsTitleShape(shape))
                    {
                        titleShape = shape;
                    }

                    try
                    {
                        containsDataVisual = containsDataVisual ||
                            shape.HasTable == Office.MsoTriState.msoTrue ||
                            shape.HasChart == Office.MsoTriState.msoTrue;
                    }
                    catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "ConsistencyScanner.Scan.Shape"); }

                    if (shape.HasTextFrame != Office.MsoTriState.msoTrue ||
                        shape.TextFrame.HasText != Office.MsoTriState.msoTrue)
                    {
                        continue;
                    }

                    string text = shape.TextFrame.TextRange.Text ?? string.Empty;
                    sourceFound = sourceFound || Regex.IsMatch(
                        text,
                        @"^\s*(source|sources|quelle|quellen)\s*[:：]",
                        RegexOptions.IgnoreCase | RegexOptions.Multiline);

                    AddWordingFindings(slide, shape, text, errors);

                    if (int.TryParse(text.Trim(), out int num) && num == slide.SlideNumber)
                    {
                        slideNumberFound = true;
                    }
                }

                if (titleShape == null)
                {
                    titleShape = FindTitleCandidate(slide);
                }
                AddTitleFindings(slide, titleShape, errors);

                if (containsDataVisual &&
                    !sourceFound &&
                    !StyleCheckRuleIgnore.IsIgnored(slide, "Consistency", "SourceLine"))
                {
                    errors.Add(new ConsistencyError
                    {
                        RuleId = "SourceLine",
                        ErrorType = "Source",
                        Description = "Chart or table has no source line",
                        SlideNumber = slide.SlideNumber,
                        ShapeName = "Slide"
                    });
                }

                if (footerExpected &&
                    slide.SlideNumber > 1 &&
                    !HasVisibleFooter(slide) &&
                    !StyleCheckRuleIgnore.IsIgnored(slide, "Consistency", "Footer"))
                {
                    errors.Add(new ConsistencyError
                    {
                        RuleId = "Footer",
                        ErrorType = "Footer",
                        Description = "Footer is hidden while other checked slides show it",
                        SlideNumber = slide.SlideNumber,
                        ShapeName = "Slide"
                    });
                }

                if (slide.SlideNumber > 1 &&
                    !slideNumberFound &&
                    !HasVisibleSlideNumber(slide) &&
                    !StyleCheckRuleIgnore.IsIgnored(slide, "Consistency", "Pagination"))
                {
                    errors.Add(new ConsistencyError
                    {
                        RuleId = "Pagination",
                        ErrorType = "Pagination",
                        Description = "Missing slide number",
                        SlideNumber = slide.SlideNumber,
                        ShapeName = "Slide"
                    });
                }
            }

            return errors;
        }

        private static void AddWordingFindings(
            Slide slide,
            Shape shape,
            string text,
            List<ConsistencyError> errors)
        {
            if (!StyleCheckRuleIgnore.IsIgnored(shape, "Consistency", "DoubleSpaces"))
            {
                foreach (Match match in Regex.Matches(text, @"[^\S\r\n]{2,}"))
                {
                    errors.Add(new ConsistencyError
                    {
                        RuleId = "DoubleSpaces",
                        ErrorType = "Wording",
                        Description = "Double spaces found",
                        SlideNumber = slide.SlideNumber,
                        ShapeName = shape.Name,
                        SuggestedCorrection = " ",
                        TextStart = match.Index + 1,
                        TextLength = match.Length
                    });
                }
            }

            if (!StyleCheckRuleIgnore.IsIgnored(shape, "Consistency", "RepeatedWord"))
            {
                foreach (Match match in Regex.Matches(
                    text,
                    @"\b(\w+)\s+\1\b",
                    RegexOptions.IgnoreCase))
                {
                    errors.Add(new ConsistencyError
                    {
                        RuleId = "RepeatedWord",
                        ErrorType = "Wording",
                        Description = "Repeated word found",
                        SlideNumber = slide.SlideNumber,
                        ShapeName = shape.Name,
                        SuggestedCorrection = match.Groups[1].Value,
                        TextStart = match.Index + 1,
                        TextLength = match.Length
                    });
                }
            }

            if (!StyleCheckRuleIgnore.IsIgnored(shape, "Consistency", "UnmatchedBrackets"))
            {
                int openBrackets = text.Count(character => character == '(');
                int closeBrackets = text.Count(character => character == ')');
                if (openBrackets != closeBrackets)
                {
                    errors.Add(new ConsistencyError
                    {
                        RuleId = "UnmatchedBrackets",
                        ErrorType = "Wording",
                        Description = "Unmatched brackets",
                        SlideNumber = slide.SlideNumber,
                        ShapeName = shape.Name
                    });
                }
            }
        }

        private static void AddTitleFindings(
            Slide slide,
            Shape titleShape,
            List<ConsistencyError> errors)
        {
            if (slide.SlideNumber <= 1) return;

            if (titleShape == null ||
                titleShape.HasTextFrame != Office.MsoTriState.msoTrue ||
                titleShape.TextFrame.HasText != Office.MsoTriState.msoTrue ||
                string.IsNullOrWhiteSpace(titleShape.TextFrame.TextRange.Text))
            {
                if (!StyleCheckRuleIgnore.IsIgnored(slide, "Consistency", "MissingTitle"))
                {
                    errors.Add(new ConsistencyError
                    {
                        RuleId = "MissingTitle",
                        ErrorType = "Title",
                        Description = "Missing slide title",
                        SlideNumber = slide.SlideNumber,
                        ShapeName = "Slide"
                    });
                }
                return;
            }

            if (StyleCheckRuleIgnore.IsIgnored(titleShape, "Consistency", "TitleLength")) return;

            string titleText = titleShape.TextFrame.TextRange.Text.Trim();
            if (titleText.Length > RecommendedTitleLength)
            {
                errors.Add(new ConsistencyError
                {
                    RuleId = "TitleLength",
                    ErrorType = "Title",
                    Description = $"Title exceeds {RecommendedTitleLength} characters",
                    SlideNumber = slide.SlideNumber,
                    ShapeName = titleShape.Name
                });
            }

            if (titleText.IndexOfAny(new[] { '\r', '\n' }) >= 0)
            {
                errors.Add(new ConsistencyError
                {
                    RuleId = "TitleLength",
                    ErrorType = "Title",
                    Description = "Title spans multiple paragraphs",
                    SlideNumber = slide.SlideNumber,
                    ShapeName = titleShape.Name
                });
            }
        }

        private static bool IsTitleShape(Shape shape)
        {
            try
            {
                if (shape.Type != Office.MsoShapeType.msoPlaceholder) return false;
                var placeholderType = shape.PlaceholderFormat.Type;
                return placeholderType == PpPlaceholderType.ppPlaceholderTitle ||
                       placeholderType == PpPlaceholderType.ppPlaceholderCenterTitle;
            }
            catch
            {
                return false;
            }
        }

        private static Shape FindTitleCandidate(Slide slide)
        {
            try
            {
                var nativeTitle = slide.Shapes.Title;
                if (nativeTitle != null) return nativeTitle;
            }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "ConsistencyScanner.GetTextStyle"); }

            Shape bestCandidate = null;
            float bestScore = float.MinValue;
            float slideHeight = 540f;
            try
            {
                slideHeight = ((Presentation)slide.Parent).PageSetup.SlideHeight;
            }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "ConsistencyScanner.GetShapeStyle"); }

            foreach (Shape shape in EnumerateShapes(slide))
            {
                try
                {
                    if (shape.HasTextFrame != Office.MsoTriState.msoTrue ||
                        shape.TextFrame.HasText != Office.MsoTriState.msoTrue ||
                        shape.Top > slideHeight * 0.35f ||
                        shape.HasTable == Office.MsoTriState.msoTrue ||
                        shape.HasChart == Office.MsoTriState.msoTrue)
                    {
                        continue;
                    }

                    string text = (shape.TextFrame.TextRange.Text ?? string.Empty).Trim();
                    if (text.Length < 5 ||
                        Regex.IsMatch(text, @"^\s*(source|sources|quelle|quellen)\s*[:：]", RegexOptions.IgnoreCase))
                    {
                        continue;
                    }

                    float fontSize = shape.TextFrame.TextRange.Font.Size;
                    float score = Math.Max(0f, fontSize) * 10f - shape.Top;
                    if (score > bestScore)
                    {
                        bestScore = score;
                        bestCandidate = shape;
                    }
                }
                catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "ConsistencyScanner.EnumerateGroup.Children"); }
            }

            return bestCandidate;
        }

        private static IEnumerable<Shape> EnumerateShapes(Slide slide)
        {
            foreach (Shape shape in slide.Shapes)
            {
                foreach (Shape item in EnumerateShapeTree(shape))
                {
                    yield return item;
                }
            }
        }

        private static IEnumerable<Shape> EnumerateShapeTree(Shape shape)
        {
            yield return shape;

            Office.MsoShapeType shapeType = Office.MsoShapeType.msoAutoShape;
            try { shapeType = shape.Type; }
            catch (Exception ex) { ExceptionPolicy.ReportOptionalComFailure(ex, "ConsistencyScanner.EnumerateGroup.ShapeType"); }
            if (shapeType != Office.MsoShapeType.msoGroup) yield break;

            foreach (Shape child in shape.GroupItems)
            {
                foreach (Shape item in EnumerateShapeTree(child))
                {
                    yield return item;
                }
            }
        }

        private static bool IsFooterExpected(Application app, List<Slide> fallbackSlides)
        {
            try
            {
                foreach (Slide slide in app.ActivePresentation.Slides)
                {
                    if (HasVisibleFooter(slide)) return true;
                }
                return false;
            }
            catch
            {
                return fallbackSlides.Any(HasVisibleFooter);
            }
        }

        private static bool HasVisibleFooter(Slide slide)
        {
            try
            {
                return slide.HeadersFooters.Footer.Visible == Office.MsoTriState.msoTrue;
            }
            catch
            {
                return false;
            }
        }

        private static bool HasVisibleSlideNumber(Slide slide)
        {
            try
            {
                return slide.HeadersFooters.SlideNumber.Visible == Office.MsoTriState.msoTrue;
            }
            catch
            {
                return false;
            }
        }
    }
}
