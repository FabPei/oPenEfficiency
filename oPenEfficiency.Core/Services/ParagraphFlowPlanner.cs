using System;
using System.Collections.Generic;
using System.Linq;

namespace oPenEfficiency.Services
{
    /// <summary>
    /// Host-free paragraph-flow rules behind Split Text Box by Paragraphs (BtnSplitByParagraphs, including its
    /// Merge Selected Text Boxes action) and Swap Text (Plain) (BtnSwapTextPlain): reading order for merging text
    /// boxes, the empty-paragraph policy and paragraph-ending handling of the split, and the paragraph labels the
    /// swap dialog shows. The commands call this planner, so the acceptance scenarios drive the production decision
    /// instead of a second test-only implementation.
    /// </summary>
    public static class ParagraphFlowPlanner
    {
        /// <summary>Shapes whose top edges differ by at most this many points count as one row (left edge decides).</summary>
        public const float ReadingOrderRowTolerancePoints = 5f;

        /// <summary>Paragraph-choice labels are truncated to this many characters (with an ellipsis).</summary>
        public const int ChoiceLabelMaximumLength = 100;

        /// <summary>Reading order for merging text boxes: top to bottom, and left to right within one row.</summary>
        public static int CompareReadingOrder(float firstTop, float firstLeft, float secondTop, float secondLeft)
        {
            if (Math.Abs(firstTop - secondTop) <= ReadingOrderRowTolerancePoints)
                return firstLeft.CompareTo(secondLeft);
            return firstTop.CompareTo(secondTop);
        }

        /// <summary>The paragraph text without its trailing paragraph marks.</summary>
        public static string WithoutParagraphEnding(string text)
        {
            return (text ?? string.Empty).TrimEnd('\r', '\n');
        }

        /// <summary>The trailing paragraph marks of the text, kept when the paragraph is swapped.</summary>
        public static string ParagraphEnding(string text)
        {
            if (string.IsNullOrEmpty(text)) return string.Empty;
            int start = text.Length;
            while (start > 0 && (text[start - 1] == '\r' || text[start - 1] == '\n')) start--;
            return text.Substring(start);
        }

        /// <summary>Splitting keeps every paragraph when asked for; otherwise only paragraphs with visible text.</summary>
        public static bool KeepParagraphForSplit(string text, bool keepEmptyParagraphs)
        {
            return keepEmptyParagraphs || !string.IsNullOrWhiteSpace(text);
        }

        /// <summary>One merged text box: the texts in reading order, each stripped of its trailing paragraph marks.</summary>
        public static string JoinParagraphs(IEnumerable<string> texts)
        {
            return string.Join("\r", (texts ?? Enumerable.Empty<string>()).Select(WithoutParagraphEnding));
        }

        /// <summary>The label of one paragraph choice: index, text, "(empty paragraph)" for blanks, truncated.</summary>
        public static string DescribeParagraphChoice(int oneBasedIndex, string text)
        {
            string description = WithoutParagraphEnding(text)
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Trim();
            if (description.Length == 0) description = "(empty paragraph)";
            if (description.Length > ChoiceLabelMaximumLength) description = description.Substring(0, 97) + "...";
            return string.Format("{0}: {1}", oneBasedIndex, description);
        }
    }
}
