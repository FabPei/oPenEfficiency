using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace oPenEfficiency.Services
{
    public class ConsistencyError
    {
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
        public static List<ConsistencyError> ScanSlides(Application app, List<Slide> slides)
        {
            var errors = new List<ConsistencyError>();
            if (slides == null || slides.Count == 0) return errors;

            HashSet<int> existingSlideNumbers = new HashSet<int>();

            foreach (var slide in slides)
            {
                bool slideNumberFound = false;

                foreach (Shape shape in slide.Shapes)
                {
                    // Skip ignored shapes
                    try {
                        if (shape.Tags["oPE_IgnoreConsistencyCheck"] == "True") continue;
                    } catch { }

                    if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                    {
                        string text = shape.TextFrame.TextRange.Text;

                        // Check double spaces
                        if (Regex.IsMatch(text, @"\s{2,}"))
                        {
                            var doubleSpaceMatches = Regex.Matches(text, @"\s{2,}");
                        foreach (Match m in doubleSpaceMatches)
                        {
                            errors.Add(new ConsistencyError { ErrorType = "Wording", Description = "Double spaces found", SlideNumber = slide.SlideNumber, ShapeName = shape.Name, SuggestedCorrection = " ", TextStart = m.Index + 1, TextLength = m.Length });
                        }
                        }

                        // Check repeated words
                        if (Regex.IsMatch(text, @"\b(\w+)\s+\1\b", RegexOptions.IgnoreCase))
                        {
                            var repeatedWordMatches = Regex.Matches(text, @"\b(\w+)\s+\b", RegexOptions.IgnoreCase);
                        foreach (Match m in repeatedWordMatches)
                        {
                            errors.Add(new ConsistencyError { ErrorType = "Wording", Description = "Repeated word found", SlideNumber = slide.SlideNumber, ShapeName = shape.Name, SuggestedCorrection = m.Groups[1].Value, TextStart = m.Index + 1, TextLength = m.Length });
                        }
                        }

                        // Check missing closing brackets
                        int openBrackets = text.Split('(').Length - 1;
                        int closeBrackets = text.Split(')').Length - 1;
                        if (openBrackets != closeBrackets)
                        {
                            errors.Add(new ConsistencyError { ErrorType = "Wording", Description = "Unmatched brackets", SlideNumber = slide.SlideNumber, ShapeName = shape.Name });
                        }

                        // Try to find slide number (heuristic: text is just a number)
                        if (int.TryParse(text.Trim(), out int num))
                        {
                            if (num == slide.SlideNumber)
                            {
                                slideNumberFound = true;
                                existingSlideNumbers.Add(num);
                            }
                        }
                    }
                }

                // If not found and not title slide (assuming slide 1 is title)
                if (slide.SlideNumber > 1 && !slideNumberFound)
                {
                    // Check native slide number setting
                    try
                    {
                        if (slide.HeadersFooters.SlideNumber.Visible == Office.MsoTriState.msoFalse)
                        {
                            errors.Add(new ConsistencyError { ErrorType = "Pagination", Description = "Missing slide number", SlideNumber = slide.SlideNumber, ShapeName = "Slide" });
                        }
                    }
                    catch { }
                }
            }

            return errors;
        }
    }
}